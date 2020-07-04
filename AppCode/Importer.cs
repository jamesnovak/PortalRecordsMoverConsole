using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.Serialization;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace PortalRecordsMover.AppCode
{
    /// <summary>
    /// Handles the import of a website data file into the Target System 
    /// </summary>
    public class Importer
    {
        private IOrganizationService _service = null;
        private ExportSettings _settings = null;

        public IOrganizationService Service { get => _service; set => _service = value; }
        public ExportSettings Settings { get => _settings; set => _settings = value; }

        public Importer(ExportSettings settings)
        {
            Settings = settings;
            Service = new CrmServiceClient(Settings.TargetConnectionString);
        }

        public void Import()
        {
            if(Directory.Exists(Settings.Config.ImportFilename))
                throw new ApplicationException($"FOLDER IMPORT NOT IMPLEMENTED!");

            if (!File.Exists(Settings.Config.ImportFilename)) {
                throw new ApplicationException($"The file {Settings.Config.ImportFilename} does not exist!");
            }

            EntityCollection entities;

            var pManager = new PluginManager(Service);
            var nManager = new NoteManager(Service);

            if (Settings.Config.DeactivateWebPagePlugins)
            {
                PortalMover.ReportProgress($"Deactivating Webpage plugins steps");
                pManager.DeactivateWebpagePlugins();
                PortalMover.ReportProgress($"Webpage plugins steps deactivated");
            }

            if (Settings.Config.RemoveJavaScriptFileRestriction && nManager.HasJsRestriction)
            {
                PortalMover.ReportProgress($"Removing JavaScript file restriction");
                nManager.RemoveRestriction();
                PortalMover.ReportProgress($"JavaScript file restriction removed");
            }

            InitializeRecordsForImport();

            DoImport();

            if (Settings.Config.DeactivateWebPagePlugins)
            {
                PortalMover.ReportProgress($"Reactivating Webpage plugins steps");
                pManager.ActivateWebpagePlugins();
                PortalMover.ReportProgress($"Webpage plugins steps activated");
            }

            if (Settings.Config.RemoveJavaScriptFileRestriction && nManager.HasJsRestriction)
            {
                PortalMover.ReportProgress($"Adding back JavaScript file restriction");
                nManager.AddRestriction();
                PortalMover.ReportProgress($"JavaScript file restriction added back");
            }

            void InitializeRecordsForImport() {

                // load the file from disk, deserialize 
                PortalMover.ReportProgress($"Deserializing ImportFileName: {Settings.Config.ImportFilename}");
                using (var reader = new StreamReader(Settings.Config.ImportFilename))
                {
                    var serializer = new DataContractSerializer(typeof(EntityCollection), new List<Type> { typeof(Entity) });
                    entities = (EntityCollection)serializer.ReadObject(reader.BaseStream);
                }

                if (entities.Entities.Count == 0) {
                    throw new ApplicationException("No enitites available to import.");
                }

                // if any of the website IDs to be imported do not maptch the IDs in the target environment
                // then map them using the settings in the configuration file.
                if (!AllWebsiteIdsMap())
                {
                    // loop on all of the website Id mappings
                    foreach (var map in Settings.Config.WebsiteIdMapping) 
                    {
                        // find all attributes for the list of entities and update the original ID to that in the mapping 
                        var attribs = entities.Entities.SelectMany(ent => ent.Attributes)
                            .Where(a => a.Value is EntityReference && ((EntityReference)a.Value).Id == map.SourceId);

                        // update each with the mapped ID
                        foreach (var attr in attribs) {
                            ((EntityReference)attr.Value).Id = map.TargetId;
                        }
                    }
                }

                // now, check again... 
                // TODO throw exception here or just ignore and remove the items?
                if (!AllWebsiteIdsMap()) {
                    throw new ApplicationException("Records selected for import do not map to websites in the target system");
                }
            }

            void DoImport()
            {
                // load metadata for this environment 
                PortalMover.ReportProgress($"Loading Metadata for current environment: {Settings.Config.TargetEnvironment}");
                Settings.AllEntities = MetadataManager.GetEntitiesList(Service).ToList();

                PortalMover.ReportProgress("Processing records for import");

                var rm = new RecordManager(Service, Settings);

                rm.ProcessRecords(entities, Settings.AllEntities);
            };

            // query the list of records being imported and the list of websites in the target system
            // return whether all of the IDs map or not.
            bool AllWebsiteIdsMap() {

                // check to see if we have any source website IDs that do not match the target environment 
                var webSitesId = entities.Entities.SelectMany(e => e.Attributes)
                        .Where(a => a.Value is EntityReference && ((EntityReference)a.Value).LogicalName == "adx_website")
                        .Select(a => ((EntityReference)a.Value).Id)
                        .Distinct()
                        .ToList();

                var targetWebSites = Service.RetrieveMultiple(new QueryExpression("adx_website") {
                    ColumnSet = new ColumnSet("adx_name")
                }).Entities;

                return webSitesId.All(id => targetWebSites.Select(w => w.Id).Contains(id));
            }
        }
    }
}
