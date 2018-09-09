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

namespace MscrmTools.PortalRecordsMover.AppCode
{
    public class Exporter
    {
        private IOrganizationService _service = null;
        private ExportSettings _settings = null;

        public IOrganizationService Service { get => _service; set => _service = value; }
        public ExportSettings Settings { get => _settings; set => _settings = value; }

        public Exporter(ExportSettings settings)
        {
            this.Settings = settings;
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
                foreach (var record in exportRecords.Entities) {
                    var emd = Settings.AllEntities.First(ent => ent.LogicalName == record.LogicalName);

                    if (!emd.IsIntersect.Value) {
                        var validAttributes = emd.Attributes
                            .Where(a => a.IsValidForCreate.Value || a.IsValidForUpdate.Value)
                            .Select(a => a.LogicalName)
                            .ToArray();

                        for (int i = record.Attributes.Count - 1; i >= 0; i--) {
                            var attr = record.Attributes.ElementAt(i);
                            if (!validAttributes.Contains(attr.Key)) {
                                // PortalMover.ReportProgress($"Export: removing attribute {attr.Key}");
                                record.Attributes.Remove(attr);
                            }
                            else {
                                var er = record[attr.Key] as EntityReference;
                                if (er != null && er.LogicalName == "contact") {
                                    // PortalMover.ReportProgress($"Export: removing attribute {attr.Key}");
                                    record.Attributes.Remove(attr);
                                }
                            }
                        }
                    }
                }
            }

            void RetrieveRecordsForExport()
            {
                // load metadata for this environment 
                PortalMover.ReportProgress($"Loading Metadata for current environment: {Settings.SourceEnvironment}");
                Settings.AllEntities = MetadataManager.GetEntitiesList(Service).ToList();

                // load records based on the date settings: entities list, filters, etc.
                var results = RetrieveRecords();

                // get the entity records from the retrieve records results 
                foreach (var entity in results.Entities) {
                    if (entity.Records.Entities.Count > 0) {
                        PortalMover.ReportProgress($"Export: Adding records for export - Entity: {entity.Records.EntityName}, Count:{entity.Records.Entities.Count}");
                        exportRecords.Entities.AddRange(entity.Records.Entities);
                    }
                }
                foreach (var entity in results.NnRecords) {
                    if (entity.Records.Entities.Count > 0) {
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
                var xwSettings = new XmlWriterSettings { Indent = true };
                var serializer = new DataContractSerializer(typeof(EntityCollection), new List<Type> { typeof(Entity) });

                using (var w = XmlWriter.Create(Settings.ExportFilename, xwSettings)) {
                    serializer.WriteObject(w, exportRecords);
                }
                PortalMover.ReportProgress($"Records exported to {Settings.ExportFilename}!");
            }
        }

        private List<Entity> RetrieveViews(List<EntityMetadata> entities)
        {
            var query = new QueryExpression("savedquery") {
                ColumnSet = new ColumnSet("returnedtypecode", "layoutxml"),
                Criteria = new FilterExpression {
                    Conditions =
                    {
                        new ConditionExpression("isquickfindquery", ConditionOperator.Equal, true),
                        new ConditionExpression("returnedtypecode", ConditionOperator.In, entities.Select(e=>e.LogicalName).ToArray())
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

            results.Views = RetrieveViews(Settings.Entities);

            // execute multiple!
            var execMulti = new ExecuteMultipleRequest() {
                Requests = new OrganizationRequestCollection(),
                Settings = new ExecuteMultipleSettings() {
                    ContinueOnError = true,
                    ReturnResponses = true
                }
            };

            var counter = 0;
            var currList = new List<string>();

            foreach (var entity in Settings.Entities) 
            {
                currList.Add($"{entity.LogicalName}: '{entity.DisplayName?.UserLocalizedLabel?.Label ?? entity.LogicalName}'");
                execMulti.Requests.Add(GetRetieveMultipleRequest(entity, Settings));
                ++counter;
                // batch in tens, or until we reach the end
                if ((currList.Count == Settings.BatchCount) || (Settings.Entities.Count == counter)) {

                    PortalMover.ReportProgress($"Begin Execute Multiple Request for entity records:\n{string.Join(", ", currList.ToArray())}");
                    var multiResults = Service.Execute(execMulti) as ExecuteMultipleResponse;

                    // iterate on results and add to return 
                    foreach (var result in multiResults.Responses) {
                        // get the restrieve multiple response and make sure we have records
                        var entities = ((RetrieveMultipleResponse)result.Response).EntityCollection;

                        if (entities.Entities.Count > 0) {
                            // PortalMover.ReportProgress($"Returned {entities.Entities.Count} records for {entities.EntityName}");
                            var er = new EntityResult { Records = entities };
                            results.Entities.Add(er);
                        }
                    }

                    // now that we have retrieved these entities, clear the collection 
                    execMulti.Requests.Clear();
                    currList.Clear();
                }
            }

            PortalMover.ReportProgress("Retrieving many to many relationships records.");
            results.NnRecords = RetrieveNnRecords(Settings, results.Entities.SelectMany(a => a.Records.Entities).ToList());

            return results;
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

            if (settings.CreateFilter.HasValue) {
                dateFilter.Filters.Add(new FilterExpression(LogicalOperator.Or));
                dateFilter.Conditions.Add(new ConditionExpression("createdon", ConditionOperator.OnOrAfter, settings.CreateFilter.Value.ToString("yyyy-MM-dd")));
            }

            if (settings.ModifyFilter.HasValue) {
                dateFilter.Filters.Add(new FilterExpression(LogicalOperator.Or));
                dateFilter.Conditions.Add(new ConditionExpression("modifiedon", ConditionOperator.OnOrAfter, settings.ModifyFilter.Value.ToString("yyyy-MM-dd")));
                // query.Criteria.AddCondition("modifiedon", ConditionOperator.OnOrAfter, settings.ModifyFilter.Value.ToString("yyyy-MM-dd"));
            }

            // add the OR date filter for createdon and modified on
            query.Criteria.Filters.Add(dateFilter);

            if (settings.WebsiteFilter != Guid.Empty && emd.Attributes.Any(a => a is LookupAttributeMetadata && ((LookupAttributeMetadata)a).Targets[0] == "adx_website")) {
                query.Criteria.AddCondition("adx_websiteid", ConditionOperator.Equal, settings.WebsiteFilter);
            }

            // add the Active check if the statecode attribute is present
            if (settings.ActiveItemsOnly && (emd.Attributes.Where(a => a.LogicalName == "statecode").ToArray().Length > 0)) {
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
            return Service.RetrieveMultiple(new QueryExpression("annotation") {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression {
                    Conditions = { new ConditionExpression("objectid", ConditionOperator.In, ids.ToArray()) }
                }
            }).Entities.ToList();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="settings"></param>
        /// <param name="records"></param>
        /// <returns></returns>
        private List<EntityResult> RetrieveNnRecords(ExportSettings settings, List<Entity> records)
        {
            var ers = new List<EntityResult>();
            var rels = new List<ManyToManyRelationshipMetadata>();

            foreach (var emd in settings.Entities) {
                foreach (var mm in emd.ManyToManyRelationships) {
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

            var execMulti = new ExecuteMultipleRequest() {
                Requests = new OrganizationRequestCollection(),
                Settings = new ExecuteMultipleSettings() {
                    ContinueOnError = true,
                    ReturnResponses = true
                }
            };

            foreach (var mm in rels) {
                var ids = records.Where(r => r.LogicalName == mm.Entity1LogicalName).Select(r => r.Id).ToList();
                if (!ids.Any()) {
                    continue;
                }

                var query = new QueryExpression(mm.IntersectEntityName) {
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
            foreach (var result in multiResults.Responses) {
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
