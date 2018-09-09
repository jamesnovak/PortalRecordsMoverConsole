using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using PortalRecordsMover.AppCode;

namespace PortalRecordsMover.AppCode
{
    internal class RecordManager
    {
        private readonly IOrganizationService service;
        private const int maxErrorLoopCount = 5;

        public RecordManager(IOrganizationService service)
        {
            this.service = service;
        }

        public bool ProcessRecords(EntityCollection ec, List<EntityMetadata> emds)
        {
            var records = new List<Entity>(ec.Entities);
            var progress = new ImportProgress(records.Count);

            var nextCycle = new List<Entity>();
            int loopIndex = 0;
            while (records.Any())
            {
                loopIndex++;
                if (loopIndex == maxErrorLoopCount)
                {
                    break;
                }

                for (int i = records.Count - 1; i >= 0; i--)
                {
                    EntityProgress entityProgress;

                    var record = records[i];
                    // Handle annotations.  
                    // TODO review this section 
                    if (record.LogicalName != "annotation")
                    {
                        // see if any records in the entity collection have a reference to this Annotation 
                        if (record.Attributes.Values.Any(v => v is EntityReference && records.Select(r => r.Id).Contains(((EntityReference)v).Id)))
                        {
                            if (nextCycle.Any(r => r.Id == record.Id)) {
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

                            foreach (var attr in toRemove) {
                                record.Attributes.Remove(attr);
                            }
                        }
                        // 
                        if (record.Attributes.Values.Any(v => (v is Guid) && records.Where(r => r.Id != record.Id).Select(r => r.Id).Contains((Guid)v))) {
                            continue;
                        }
                    }

                    // update the entity progress element 
                    TrackEntityProgress();

                    try
                    {
                        record.Attributes.Remove("ownerid");

                        // check to see if this is an N:N relation vs a standard record import.
                        if (record.Attributes.Count == 3 && record.Attributes.Values.All(v => v is Guid))
                        {
                            try
                            {
                                // perform the association!
                                var rel = emds.SelectMany(e => e.ManyToManyRelationships).First(r => r.IntersectEntityName == record.LogicalName);
                                
                                service.Associate(
                                    rel.Entity1LogicalName,
                                    record.GetAttributeValue<Guid>(rel.Entity1IntersectAttribute),
                                    new Relationship(rel.SchemaName),
                                    new EntityReferenceCollection(new List<EntityReference>
                                    {
                                        new EntityReference(rel.Entity2LogicalName, record.GetAttributeValue<Guid>(rel.Entity2IntersectAttribute))
                                    })
                                );

                                PortalMover.ReportProgress($"Import: Association {entityProgress.Entity} ({record.Id}) created");
                            }
                            catch (FaultException<OrganizationServiceFault> error)
                            {
                                if (error.Detail.ErrorCode != -2147220937) {
                                    throw;
                                }
                                PortalMover.ReportProgress($"Import: Association {entityProgress.Entity} ({record.Id}) already exists");
                            }
                        }
                        else
                        {
                            // Do the Insert/Update!
                            var result = (UpsertResponse)service.Execute(new UpsertRequest {
                                Target = record
                            });
                            PortalMover.ReportProgress($"Import: Record {record.GetAttributeValue<string>(entityProgress.Metadata.PrimaryNameAttribute)} {(result.RecordCreated ? "created" : "updated")} ({entityProgress.Entity}/{record.Id})");
                        }

                        records.RemoveAt(i);
                        entityProgress.Success++;
                        entityProgress.Processed++;
                    }
                    catch (Exception error)
                    {
                        PortalMover.ReportProgress($"Import: An error occured attempting the insert/update/associate: {record.GetAttributeValue<string>(entityProgress.Metadata.PrimaryNameAttribute)} ({entityProgress.Entity}/{record.Id}): {error.Message}");
                        entityProgress.Error++;
                    }

                    // track the progress of the current list of records being imported 
                    void TrackEntityProgress()
                    {
                        // get the entity progress for this record entity type.
                        entityProgress = progress.Entities.FirstOrDefault(e => e.LogicalName == record.LogicalName);
                        if (entityProgress == null) 
                        {
                            var emd = emds.First(e => e.LogicalName == record.LogicalName);
                            string displayName = emd.DisplayName?.UserLocalizedLabel?.Label;

                            if (displayName == null && emd.IsIntersect.Value) {
                                var rel = emds.SelectMany(ent => ent.ManyToManyRelationships)
                                .First(r => r.IntersectEntityName == emd.LogicalName);

                                displayName = $"{emds.First(ent => ent.LogicalName == rel.Entity1LogicalName).DisplayName?.UserLocalizedLabel?.Label} / {emds.First(ent => ent.LogicalName == rel.Entity2LogicalName).DisplayName?.UserLocalizedLabel?.Label}";
                            }
                            if (displayName == null) {
                                displayName = emd.SchemaName;
                            }

                            entityProgress = new EntityProgress(emd, displayName);
                            progress.Entities.Add(entityProgress);
                        }
                    }
                }
            }

            PortalMover.ReportProgress("Import: Updating records to add references");

            var count = nextCycle.DistinctBy(r => r.Id).Count();
            var index = 0;

            foreach (var record in nextCycle.DistinctBy(r => r.Id))
            {
                try
                {
                    index++;

                    PortalMover.ReportProgress($"Import: Upating record {record.LogicalName} ({record.Id})");

                    record.Attributes.Remove("ownerid");
                    service.Update(record);
                }
                catch (Exception error)
                {
                    PortalMover.ReportProgress ($"Import: An error occured during import: {error.Message}");
                }
            }
            return false;
        }
    }
}