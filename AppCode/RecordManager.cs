using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace PortalRecordsMover.AppCode
{
    internal class RecordManager
    {
        private const int maxErrorLoopCount = 5;
        private readonly List<EntityReference> recordsToDeactivate;
        private readonly IOrganizationService service = null;
        private readonly ExportSettings settings = null;

        public RecordManager(IOrganizationService service, ExportSettings settings): this(service)
        {
            this.settings = settings;
        }

        public RecordManager(IOrganizationService service)
        {
            this.service = service;
            recordsToDeactivate = new List<EntityReference>();
        }

        public bool ProcessRecords(EntityCollection ec, List<EntityMetadata> emds, int organizationMajorVersion = 9)
        {
            var records = new List<Entity>(ec.Entities);

            // Move annotation at the beginning if the list as the list will be
            // inverted to allow list removal. This way, annotation are
            // processed as the last records
            var annotations = records.Where(e => e.LogicalName == "annotation").ToList();
            records = records.Except(annotations).ToList();
            records.InsertRange(0, annotations);

            var progress = new ImportProgress(records.Count);

            var nextCycle = new List<Entity>();

            // reset Error state for each
            progress.Entities.ForEach(p => { p.ErrorFirstPhase = 0; });

            for (int i = records.Count - 1; i >= 0; i--)
            {
                EntityProgress entityProgress;

                var record = records[i];

                // Handle annotations.  
                if (record.LogicalName != "annotation")
                {
                    // see if any records in the entity collection have a reference to this Annotation 
                    if (record.Attributes.Values.Any(v =>
                        v is EntityReference reference
                        && records.Select(r => r.Id).Contains(reference.Id)
                        ))
                    {
                        if (nextCycle.Any(r => r.Id == record.Id))
                        {
                            continue;
                        }

                        var newRecord = new Entity(record.LogicalName) { Id = record.Id };
                        var toRemove = new List<string>();
                        foreach (var attr in record.Attributes)
                        {
                            if (attr.Value is EntityReference)
                            {
                                newRecord.Attributes.Add(attr.Key, attr.Value);
                                toRemove.Add(attr.Key);
                                nextCycle.Add(newRecord);
                            }
                        }

                        foreach (var attr in toRemove)
                        {
                            record.Attributes.Remove(attr);
                        }
                    }

                    if (record.Attributes.Values.Any(v =>
                        v is Guid guid
                        && records.Where(r => r.Id != record.Id)
                            .Select(r => r.Id)
                            .Contains(guid)
                            ))
                    {
                        nextCycle.Add(record);
                        records.RemoveAt(i);
                        continue;
                    }
                }
                // check for entity progress.. entity may not be in target system
                entityProgress = GetEntityProgress(records, record, progress, emds);
                if (entityProgress == null) {
                    continue;
                }
                try
                {
                    record.Attributes.Remove("ownerid");

                    if (record.Attributes.Contains("statecode") &&
                        record.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                    {
                        PortalMover.ReportProgress($"Import: Record {record.GetAttributeValue<string>(entityProgress.Metadata.PrimaryNameAttribute)} is inactive : Added for deactivation step");

                        recordsToDeactivate.Add(record.ToEntityReference());
                        record.Attributes.Remove("statecode");
                        record.Attributes.Remove("statuscode");
                    }

                    if (organizationMajorVersion >= 8)
                    {
                        var result = (UpsertResponse)service.Execute(new UpsertRequest
                        {
                            Target = record
                        });

                        PortalMover.ReportProgress($"Import: Record {record.GetAttributeValue<string>(entityProgress.Metadata.PrimaryNameAttribute)} {(result.RecordCreated ? "created" : "updated")} ({entityProgress.Entity}/{record.Id})");
                    }
                    else
                    {
                        bool exists = false;
                        try
                        {
                            service.Retrieve(record.LogicalName, record.Id, new ColumnSet());
                            exists = true;
                        }
                        catch
                        {
                            // Do nothing
                        }

                        if (exists)
                        {
                            service.Update(record);
                            PortalMover.ReportProgress($"Import: Record {record.GetAttributeValue<string>(entityProgress.Metadata.PrimaryNameAttribute)} updated ({entityProgress.Entity}/{record.Id})");
                        }
                        else
                        {
                            service.Create(record);
                            PortalMover.ReportProgress($"Import: Record {record.GetAttributeValue<string>(entityProgress.Metadata.PrimaryNameAttribute)} created ({entityProgress.Entity}/{record.Id})");
                        }
                    }

                    if (record.LogicalName == "annotation" && settings.Config.CleanWebFiles)
                    {
                        var reference = record.GetAttributeValue<EntityReference>("objectid");
                        if (reference?.LogicalName == "adx_webfile")
                        {
                            PortalMover.ReportProgress($"Import: Searching for extra annotation in web file {reference.Id:B}");

                            var qe = new QueryExpression("annotation")
                            {
                                NoLock = true,
                                Criteria = new FilterExpression
                                {
                                    Conditions =
                                            {
                                                new ConditionExpression("annotationid", ConditionOperator.NotEqual,
                                                    record.Id),
                                                new ConditionExpression("objectid", ConditionOperator.Equal,
                                                    reference.Id),
                                            }
                                }
                            };

                            var extraNotes = service.RetrieveMultiple(qe);
                            foreach (var extraNote in extraNotes.Entities)
                            {
                                PortalMover.ReportProgress($"Import: Deleting extra note {extraNote.Id:B}");
                                service.Delete(extraNote.LogicalName, extraNote.Id);
                            }
                        }
                    }

                    records.RemoveAt(i);
                    entityProgress.SuccessFirstPhase++;
                }
                catch (Exception error)
                {
                    PortalMover.ReportProgress($"Import: {record.GetAttributeValue<string>(entityProgress.Metadata.PrimaryNameAttribute)} ({entityProgress.Entity}/{record.Id}): {error.Message}");
                    entityProgress.ErrorFirstPhase++;
                }
                finally
                {
                    entityProgress.Processed++;
                }
            }

            PortalMover.ReportProgress($"Import: Updating records to add references and processing many-to-many relationships...");

            var count = nextCycle.DistinctBy(r => r.Id).Count();
            var index = 0;

            foreach (var record in nextCycle.DistinctBy(r => r.Id))
            {
                var entityProgress = GetEntityProgress(nextCycle, record, progress, emds);

                if (!entityProgress.SuccessSecondPhase.HasValue)
                {
                    entityProgress.SuccessSecondPhase = 0;
                    entityProgress.ErrorSecondPhase = 0;
                }

                try
                {
                    index++;

                    PortalMover.ReportProgress($"Import: Upating record {record.LogicalName} ({record.Id})");

                    record.Attributes.Remove("ownerid");

                    if (record.Attributes.Count == 3 && record.Attributes.Values.All(v => v is Guid))
                    {
                        PortalMover.ReportProgress($"Import: Creating association {entityProgress.Entity} ({record.Id})");

                        try
                        {
                            var rel =
                                emds.SelectMany(e => e.ManyToManyRelationships)
                                    .First(r => r.IntersectEntityName == record.LogicalName);

                            service.Associate(rel.Entity1LogicalName,
                                record.GetAttributeValue<Guid>(rel.Entity1IntersectAttribute),
                                                    new Relationship(rel.SchemaName),
                                                    new EntityReferenceCollection(new List<EntityReference>
                                                    {
                                    new EntityReference(rel.Entity2LogicalName,
                                        record.GetAttributeValue<Guid>(rel.Entity2IntersectAttribute))
                                                    }));

                            PortalMover.ReportProgress($"Import: Association {entityProgress.Entity} ({record.Id}) created");
                        }
                        catch (FaultException<OrganizationServiceFault> error)
                        {
                            if (error.Detail.ErrorCode != -2147220937)
                            {
                                throw;
                            }

                            PortalMover.ReportProgress($"Import: Association {entityProgress.Entity} ({record.Id}) already exists");
                        }
                        finally
                        {
                            entityProgress.Processed++;
                        }
                    }
                    else
                    {
                        service.Update(record);
                    }
                }
                catch (Exception error)
                {
                    PortalMover.ReportProgress($"Import: An error occured during import: {error.Message}");
                    entityProgress.ErrorSecondPhase = entityProgress.ErrorSecondPhase.Value + 1;
                }
            }

            #region Deactivate Records
            if (recordsToDeactivate.Any())
            {
                count = recordsToDeactivate.Count;
                index = 0;

                PortalMover.ReportProgress($"Import: Deactivating records...");

                foreach (var er in recordsToDeactivate)
                {
                    var entityProgress = progress.Entities.First(e => e.LogicalName == er.LogicalName);
                    if (!entityProgress.SuccessSecondPhase.HasValue)
                    {
                        entityProgress.SuccessSetStatePhase = 0;
                        entityProgress.ErrorSetState = 0;
                    }

                    try
                    {
                        index++;

                        PortalMover.ReportProgress($"Import: Deactivating record {er.LogicalName} ({er.Id})");

                        var recordToUpdate = new Entity(er.LogicalName)
                        {
                            Id = er.Id,
                            ["statecode"] = new OptionSetValue(1),
                            ["statuscode"] = new OptionSetValue(-1)
                        };

                        service.Update(recordToUpdate);

                        var percentage = index * 100 / count;
                        entityProgress.SuccessSetStatePhase++;
                    }
                    catch (Exception error)
                    {
                        var percentage = index * 100 / count;
                        entityProgress.ErrorSetState++;
                    }
                }
            }
            #endregion

            return false;
        }

        // helper function to track the progress of the current list of records being imported 
        private EntityProgress GetEntityProgress(List<Entity> coll, Entity record, ImportProgress progress, List<EntityMetadata> emds)
        {
            // get the entity progress for this record entity type.
            var entityProgress = progress.Entities.FirstOrDefault(e => e.LogicalName == record.LogicalName);

            if (entityProgress == null)
            {
                var emd = emds.FirstOrDefault(e => e.LogicalName == record.LogicalName);
                if (emd == null)
                {
                    PortalMover.ReportProgress($"Record: Entity Logical Name: {record.LogicalName} for ID: {record.Id} not found in the target instance metadata.");
                    return null;
                }
                string displayName = emd.DisplayName?.UserLocalizedLabel?.Label;

                if (displayName == null && emd.IsIntersect.Value)
                {
                    var rel = emds.SelectMany(ent => ent.ManyToManyRelationships)
                    .First(r => r.IntersectEntityName == emd.LogicalName);
                    var nameOne = emds.FirstOrDefault(ent => ent.LogicalName == rel.Entity1LogicalName)?.DisplayName?.UserLocalizedLabel?.Label;
                    var nameTwo = emds.FirstOrDefault(ent => ent.LogicalName == rel.Entity2LogicalName)?.DisplayName?.UserLocalizedLabel?.Label;
                    displayName = $"{nameOne } / {nameTwo}";
                }
                if (displayName == null)
                {
                    displayName = emd.SchemaName;
                }

                entityProgress = new EntityProgress(emd, displayName)
                {
                    Count = coll.Count(r => r.LogicalName == record.LogicalName)
                };
                progress.Entities.Add(entityProgress);
            }
            return entityProgress;
        }

        public List<EntityResult> RetrieveNnRecords(ExportSettings exSettings, List<Entity> records)
        {
            var ers = new List<EntityResult>();
            var rels = new List<ManyToManyRelationshipMetadata>();

            foreach (var emd in exSettings.Entities)
            {
                foreach (var mm in emd.ManyToManyRelationships)
                {
                    var e1 = mm.Entity1LogicalName;
                    var e2 = mm.Entity2LogicalName;
                    var isValid = false;

                    if (e1 == emd.LogicalName)
                    {
                        if (exSettings.Entities.Any(e => e.LogicalName == e2))
                        {
                            isValid = true;
                        }
                    }
                    else
                    {
                        if (exSettings.Entities.Any(e => e.LogicalName == e1))
                        {
                            isValid = true;
                        }
                    }

                    if (isValid && rels.All(r => r.IntersectEntityName != mm.IntersectEntityName))
                    {
                        rels.Add(mm);
                    }
                }
            }

            foreach (var mm in rels)
            {
                var ids = records.Where(r => r.LogicalName == mm.Entity1LogicalName).Select(r => r.Id).ToList();
                if (!ids.Any())
                {
                    continue;
                }

                var query = new QueryExpression(mm.IntersectEntityName)
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression(mm.Entity1IntersectAttribute, ConditionOperator.In, ids.ToArray())
                        }
                    }
                };

                ers.Add(new EntityResult { Records = service.RetrieveMultiple(query) });
            }

            return ers;
        }
    }
}