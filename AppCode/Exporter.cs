using System;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Xml;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Messages;
using System.IO;
using System.Xml.Linq;

namespace PortalRecordsMover.AppCode
{
    public class Exporter
    {
        private IOrganizationService _service = null;
        private ExportSettings _settings = null;

        public IOrganizationService Service { get => _service; set => _service = value; }
        public ExportSettings Settings { get => _settings; set => _settings = value; }

        public Exporter(ExportSettings settings)
        {
            Settings = settings;
            Service = new CrmServiceClient(Settings.SourceConnectionString);
        }

        public void Export()
        {
            // kick off the async method that will retrieve the records
            var exportRecords = new EntityCollection();

            RetrieveRecordsForExport();

            PrepareAttributes();
            
            SaveToFile();

            // clean up any attributes that might not be available for import 
            void PrepareAttributes()
            {
                foreach (var record in exportRecords.Entities) 
                {
                    var emd = Settings.AllEntities.First(ent => ent.LogicalName == record.LogicalName);

                    if (!emd.IsIntersect.Value) {
                        var validAttributes = emd.Attributes
                            .Where(a => a.IsValidForCreate.Value || a.IsValidForUpdate.Value)
                            .Select(a => a.LogicalName)
                            .ToArray();

                        for (int i = record.Attributes.Count - 1; i >= 0; i--) 
                        {
                            var attr = record.Attributes.ElementAt(i);
                            if (!validAttributes.Contains(attr.Key)) 
                            {
                                // PortalMover.ReportProgress($"Export: removing attribute {attr.Key}");
                                record.Attributes.Remove(attr);
                            }
                            else 
                            {
                                var er = record[attr.Key] as EntityReference;
                                if (er != null && er.LogicalName == "contact") 
                                {
                                    // PortalMover.ReportProgress($"Export: removing attribute {attr.Key}");
                                    record.Attributes.Remove(attr);
                                }
                            }
                        }

                        foreach (var va in validAttributes)
                        {
                            //add any null attributes to force them to update to null.
                            if (!record.Contains(va))
                                record[va] = null;
                        }
                    }
                }
            }
            

            void RetrieveRecordsForExport()
            {
                // load metadata for this environment 
                PortalMover.ReportProgress($"Loading Metadata for current environment: {Settings.Config.SourceEnvironment}");
                // exclude intersect records here
                Settings.AllEntities = MetadataManager
                    .GetEntitiesList(Service)
                    .ToList();

                // load records based on the date settings: entities list, filters, etc.
                var results = RetrieveRecords();

                // get the entity records from the retrieve records results 
                foreach (var entity in results.Entities) 
                {
                    if (entity.Records.Entities.Count > 0) 
                    {
                        PortalMover.ReportProgress($"Export: Adding records for export - Entity: {entity.Records.EntityName}, Count:{entity.Records.Entities.Count}");
                        exportRecords.Entities.AddRange(entity.Records.Entities);
                    }
                }
                foreach (var entity in results.NnRecords) 
                {
                    if (entity.Records.Entities.Count > 0) 
                    {
                        PortalMover.ReportProgress($"Adding N:N records for export - Entity: {entity.Records.EntityName}, Count:{entity.Records.Entities.Count}");
                        exportRecords.Entities.AddRange(entity.Records.Entities);
                    }
                }

                // grab web files, if any are selected 
                var webFiles = exportRecords.Entities.Where(ent => ent.LogicalName == "adx_webfile").ToList();
                if (webFiles.Any()) {
                    PortalMover.ReportProgress("Retrieving web files annotation records");

                    var annotations = RetrieveWebfileAnnotations(webFiles.Select(w => w.Id).ToList());

                    foreach (var note in annotations) {
                        exportRecords.Entities.Insert(0, note);
                    }
                }
            }
            void SaveToFile()
            {
                if (Settings.Config.ExportInFolderStructure)
                {
                    ExportToFolderStructure(exportRecords);
                    return;
                }


                var xwSettings = new XmlWriterSettings { Indent = true };
                var serializer = new DataContractSerializer(typeof(EntityCollection), new List<Type> { typeof(Entity) });

                using (var w = XmlWriter.Create(Settings.Config.ExportFilename, xwSettings)) {
                    serializer.WriteObject(w, exportRecords);
                }
                PortalMover.ReportProgress($"Records exported to {Settings.Config.ExportFilename}!");
            }

            void ExportToFolderStructure(EntityCollection list)
            {
                var rootPath = string.Format(Settings.Config.ExportFilename, DateTime.Now);
                if (!Directory.Exists(rootPath))
                {
                    Directory.CreateDirectory(rootPath);
                }
                var entities = list.Entities.GroupBy(ent => ent.LogicalName);
                foreach (var entity in entities)
                {
                    var directory = Path.Combine(rootPath, entity.Key);
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    foreach (var record in entity)
                    {
                        var filePath = Path.Combine(directory, $"{record.Id:B}.xml");
                        var xwSettings = new XmlWriterSettings { Indent = true };
                        var serializer = new DataContractSerializer(typeof(Entity));

                        using (var w = XmlWriter.Create(filePath, xwSettings))
                        {
                            serializer.WriteObject(w, record);
                        }

                        CleanAndOrderXml(filePath);
                    }
                }
            }

        }
        private void CleanAndOrderXml(string filePath)
        {
            var xDoc = XDocument.Load(filePath);
            if (xDoc.Root == null) return;

            XNamespace ns = xDoc.Root.GetDefaultNamespace();

            // Order attributes
            var attributesNode = xDoc.Root.Element(ns + "Attributes");
            if (attributesNode != null)
            {
                var kvps = attributesNode
                    .Elements()
                    .OrderBy(s => ((XElement)s.FirstNode).Value)
                    .ToList();

                attributesNode.RemoveNodes();
                attributesNode.Add(kvps);
            }

            if (false)  // Remove Formatted Values selected
            {
                var formattedValues = xDoc.Root.Element(ns + "FormattedValues");
                formattedValues?.Remove();
            }

            xDoc.Save(filePath);
        }

        public List<Entity> RetrieveViews(List<EntityMetadata> entities)
        {
            var query = new QueryExpression("savedquery")
            {
                ColumnSet = new ColumnSet("returnedtypecode", "layoutxml"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("isquickfindquery", ConditionOperator.Equal, true),
                        new ConditionExpression("returnedtypecode", ConditionOperator.In, entities.Select(e=>e.LogicalName).Cast<object>().ToArray())
                    }
                }
            };

            return Service.RetrieveMultiple(query).Entities.ToList();
        }

        /// <summary>
        /// Retrieve all records for the selected entities
        /// </summary>
        /// <returns></returns>
        private ExportResults RetrieveRecords()
        {
            var results = new ExportResults { Settings = Settings };
            PortalMover.ReportProgress("Retrieving selected entities views layout");
            var entities = Settings.Entities.Where(m => m.IsIntersect == null || m.IsIntersect.Value == false);
            foreach (var entity in entities)
            {
                PortalMover.ReportProgress($"Retrieving records for entity {entity.DisplayName?.UserLocalizedLabel?.Label ?? entity.LogicalName}...");

                var er = new EntityResult
                {
                    Records = RetrieveRecords(entity, Settings.Config)
                };

                results.Entities.Add(er);
            }

            PortalMover.ReportProgress("Retrieving many to many relationships records.");
            results.NnRecords = RetrieveNnRecords(Settings, results.Entities.SelectMany(a => a.Records.Entities).ToList());

            return results;
        }

        public EntityCollection RetrieveRecords(EntityMetadata emd, MoverSettingsConfig settings)
        {
            var query = new QueryExpression(emd.LogicalName)
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression { Filters = { new FilterExpression(LogicalOperator.Or) } }
            };

            // filter by either the created on OR the modified on 
            var dateFilter = new FilterExpression(LogicalOperator.Or);

            if (settings.CreateFilter.HasValue && !emd.IsIntersect.Value)
            {
                dateFilter.Filters.Add(new FilterExpression(LogicalOperator.Or));
                dateFilter.Conditions.Add(new ConditionExpression("createdon", ConditionOperator.OnOrAfter, settings.CreateFilter.Value.ToString("yyyy-MM-dd")));
            }

            if (settings.ModifyFilter.HasValue && !emd.IsIntersect.Value)
            {
                dateFilter.Filters.Add(new FilterExpression(LogicalOperator.Or));
                dateFilter.Conditions.Add(new ConditionExpression("modifiedon", ConditionOperator.OnOrAfter, settings.ModifyFilter.Value.ToString("yyyy-MM-dd")));
            }

            // add the OR date filter for createdon and modified on
            query.Criteria.Filters.Add(dateFilter);

            if (settings.WebsiteFilter != Guid.Empty)
            {
                var lamd = emd.Attributes.FirstOrDefault(a =>
                    a is LookupAttributeMetadata metadata && 
                    metadata.Targets.Length > 0 &&
                    metadata.Targets[0] == "adx_website");

                if (lamd != null)
                {
                    query.Criteria.AddCondition(lamd.LogicalName, ConditionOperator.Equal, settings.WebsiteFilter);
                }
                else
                {
                    switch (emd.LogicalName)
                    {
                        case "adx_webfile":
                            var noteLe = new LinkEntity
                            {
                                LinkFromEntityName = "adx_webfile",
                                LinkFromAttributeName = "adx_webfileid",
                                LinkToAttributeName = "objectid",
                                LinkToEntityName = "annotation",
                                LinkCriteria = new FilterExpression(LogicalOperator.Or)
                            };

                            bool addLinkEntity = false;

                            if (settings.CreateFilter.HasValue)
                            {
                                noteLe.LinkCriteria.AddCondition("createdon", ConditionOperator.OnOrAfter, settings.CreateFilter.Value.ToString("yyyy-MM-dd"));
                                addLinkEntity = true;
                            }

                            if (settings.ModifyFilter.HasValue)
                            {
                                noteLe.LinkCriteria.AddCondition("modifiedon", ConditionOperator.OnOrAfter, settings.ModifyFilter.Value.ToString("yyyy-MM-dd"));
                                addLinkEntity = true;
                            }

                            if (addLinkEntity)
                            {
                                query.LinkEntities.Add(noteLe);
                            }
                            break;

                        case "adx_entityformmetadata":
                            query.LinkEntities.Add(
                                CreateParentEntityLinkToWebsite(
                                    emd.LogicalName,
                                    "adx_entityform",
                                    "adx_entityformid",
                                    "adx_entityform",
                                    settings.WebsiteFilter));
                            break;

                        case "adx_webformmetadata":
                            var le = CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_webformstep",
                                "adx_webformstepid",
                                "adx_webformstep",
                                Guid.Empty);

                            le.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                "adx_webformstep",
                                "adx_webform",
                                "adx_webformid",
                                "adx_webform",
                                settings.WebsiteFilter));

                            query.LinkEntities.Add(le);
                            break;

                        case "adx_weblink":
                            query.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_weblinksetid",
                                "adx_weblinksetid",
                                "adx_weblinkset",
                                settings.WebsiteFilter));
                            break;

                        case "adx_blogpost":
                            query.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_blogid",
                                "adx_blogid",
                                "adx_blog",
                                settings.WebsiteFilter));
                            break;

                        case "adx_communityforumaccesspermission":
                        case "adx_communityforumannouncement":
                        case "adx_communityforumthread":
                            query.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_forumid",
                                "adx_communityforumid",
                                "adx_communityforum",
                                settings.WebsiteFilter));
                            break;

                        case "adx_communityforumpost":
                            var lef = CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_forumthreadid",
                                "adx_communityforumthreadid",
                                "adx_communityforumthread",
                                Guid.Empty);

                            lef.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                "adx_communityforumthread",
                                "adx_forumid",
                                "adx_communityforumid",
                                "adx_communityforum",
                                settings.WebsiteFilter));

                            query.LinkEntities.Add(lef);

                            break;

                        case "adx_idea":
                            query.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_ideaforumid",
                                "adx_ideaforumid",
                                "adx_ideaforum",
                                settings.WebsiteFilter));
                            break;

                        case "adx_pagealert":
                        case "adx_webpagehistory":
                        case "adx_webpagelog":
                            query.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_webpageid",
                                "adx_webpageid",
                                "adx_webpage",
                                settings.WebsiteFilter));
                            break;

                        case "adx_pollsubmission":
                            query.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_pollid",
                                "adx_pollid",
                                "adx_poll",
                                settings.WebsiteFilter));
                            break;

                        case "adx_webfilelog":
                            query.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_webfileid",
                                "adx_webfileid",
                                "adx_webfile",
                                settings.WebsiteFilter));
                            break;

                        case "adx_webformsession":
                        case "adx_webformstep":
                            query.LinkEntities.Add(CreateParentEntityLinkToWebsite(
                                emd.LogicalName,
                                "adx_webform",
                                "adx_webformid",
                                "adx_webform",
                                settings.WebsiteFilter));
                            break;
                    }
                }
            }

            if (settings.ActiveItemsOnly && emd.LogicalName != "annotation")
            {
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            }

            return Service.RetrieveMultiple(query);
        }

        private LinkEntity CreateParentEntityLinkToWebsite(string fromEntity, string fromAttribute, string toAttribute, string toEntity, Guid websiteId)
        {
            var le = new LinkEntity
            {
                LinkFromEntityName = fromEntity,
                LinkFromAttributeName = fromAttribute,
                LinkToAttributeName = toAttribute,
                LinkToEntityName = toEntity,
            };

            if (websiteId != Guid.Empty)
            {
                le.LinkCriteria.AddCondition("adx_websiteid", ConditionOperator.Equal, websiteId);
            }

            return le;
        }
        /// <summary>
        /// Retrieve related records for the selected Entity
        /// </summary>
        /// <param name="emd"></param>
        /// <param name="settings"></param>
        /// <returns></returns>
        private EntityCollection RetrieveRecordsForEntity(EntityMetadata emd, ExportSettings settings)
        {
            var req = GetRetieveMultipleRequest(emd, settings);
            var resp = Service.Execute(req) as RetrieveMultipleResponse;

            return resp.EntityCollection;
        }

        /// <summary>
        /// Helper method that will create the retrieve multiple request for the current filter and an entity.
        /// </summary>
        /// <param name="emd"></param>
        /// <param name="settings"></param>
        /// <returns></returns>
        private RetrieveMultipleRequest GetRetieveMultipleRequest(EntityMetadata emd, ExportSettings settings)
        {
            var query = new QueryExpression(emd.LogicalName) {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression()
            };

            // filter by either the created on OR the modified on 
            var dateFilter = new FilterExpression(LogicalOperator.Or);

            if (settings.Config.CreateFilter.HasValue) {
                dateFilter.Filters.Add(new FilterExpression(LogicalOperator.Or));
                dateFilter.Conditions.Add(new ConditionExpression("createdon", ConditionOperator.OnOrAfter, settings.Config.CreateFilter.Value.ToString("yyyy-MM-dd")));
            }

            if (settings.Config.ModifyFilter.HasValue) {
                dateFilter.Filters.Add(new FilterExpression(LogicalOperator.Or));
                dateFilter.Conditions.Add(new ConditionExpression("modifiedon", ConditionOperator.OnOrAfter, settings.Config.ModifyFilter.Value.ToString("yyyy-MM-dd")));
            }

            // add the OR date filter for createdon and modified on
            query.Criteria.Filters.Add(dateFilter);

            if (settings.Config.WebsiteFilter != Guid.Empty && emd.Attributes.Any(a => a is LookupAttributeMetadata && ((LookupAttributeMetadata)a).Targets[0] == "adx_website")) {
                query.Criteria.AddCondition("adx_websiteid", ConditionOperator.Equal, settings.Config.WebsiteFilter);
            }

            // add the Active check if the statecode attribute is present
            if (settings.Config.ActiveItemsOnly && (emd.Attributes.Where(a => a.LogicalName == "statecode").ToArray().Length > 0)) {
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            }

            return new RetrieveMultipleRequest() { Query = query };
        }

        /// <summary>
        /// Retrieve all web file notifications 
        /// </summary>
        /// <param name="ids"></param>
        /// <returns></returns>
        private List<Entity> RetrieveWebfileAnnotations(List<Guid> ids)
        {
            PortalMover.ReportProgress($"Retrieving Web File Annotations for {ids.Count} records");
            return Service.RetrieveMultiple(new QueryExpression("annotation") 
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression 
                {
                    Conditions = { 
                        new ConditionExpression("objectid", ConditionOperator.In, ids.ToArray()) 
                    }
                }
            }).Entities.ToList();
        }

        /// <summary>
        /// Retrieve records for the N:N relations
        /// </summary>
        /// <param name="settings"></param>
        /// <param name="records"></param>
        /// <returns></returns>
        private List<EntityResult> RetrieveNnRecords(ExportSettings settings, List<Entity> records)
        {
            var ers = new List<EntityResult>();
            var rels = new List<ManyToManyRelationshipMetadata>();

            foreach (var emd in settings.Entities) 
            {
                foreach (var mm in emd.ManyToManyRelationships) 
                {
                    var e1 = mm.Entity1LogicalName;
                    var e2 = mm.Entity2LogicalName;
                    var isValid = false;

                    if (e1 == emd.LogicalName) {
                        if (settings.Entities.Any(e => e.LogicalName == e2)) {
                            isValid = true;
                        }
                    }
                    else {
                        if (settings.Entities.Any(e => e.LogicalName == e1)) {
                            isValid = true;
                        }
                    }

                    if (isValid && rels.All(r => r.IntersectEntityName != mm.IntersectEntityName)) {
                        rels.Add(mm);
                    }
                }
            }

            var execMulti = new ExecuteMultipleRequest() 
            {
                Requests = new OrganizationRequestCollection(),
                Settings = new ExecuteMultipleSettings() 
                {
                    ContinueOnError = true,
                    ReturnResponses = true
                }
            };

            foreach (var mm in rels) {
                var ids = records.Where(r => r.LogicalName == mm.Entity1LogicalName)
                    .Select(r => r.Id)
                    .ToList();

                if (!ids.Any()) {
                    continue;
                }

                var query = new QueryExpression(mm.IntersectEntityName) 
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression {
                        Conditions = {
                            new ConditionExpression(mm.Entity1IntersectAttribute, ConditionOperator.In, ids.ToArray())
                        }
                    }
                };
                execMulti.Requests.Add(new RetrieveMultipleRequest() { Query = query });
            }

            var multiResults = Service.Execute(execMulti) as ExecuteMultipleResponse;

            // iterate on results and add to return 
            foreach (var result in multiResults.Responses) 
            {
                // get the restrieve multiple response and make sure we have records
                var entities = ((RetrieveMultipleResponse)result.Response).EntityCollection;
                if (entities.Entities.Count > 0) {

                    var er = new EntityResult { Records = entities };
                    PortalMover.ReportProgress($"New N:N records retrieved for {er.Records.EntityName}, Count:{er.Records.Entities.Count}");
                    ers.Add(er);
                }
            }

            return ers;
        }
    }
}
