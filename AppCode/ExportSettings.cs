using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Microsoft.Xrm.Sdk.Metadata;

namespace PortalRecordsMover.AppCode
{
    /// <summary>
    /// General export settings class
    /// </summary>
    public class ExportSettings
    {
        public ExportSettings()
        {
            SettingsFileName = "ExportSettings.json";
        }

        #region Properties 
        public string SettingsFileName { get; set; }

        private MoverSettingsConfig config;
        public MoverSettingsConfig Config { get => config; set => config = value; }

        public string SourceConnectionString {
            get { return $"RequireNewInstance=True;AuthType=Office365;Username={config.SourceUsername}; Password={config.SourcePassword};Url={config.SourceEnvironment}"; }
        }
        public string TargetConnectionString {
            get { return $"RequireNewInstance=True;AuthType=Office365;Username={config.TargetUsername}; Password={config.TargetPassword};Url={config.TargetEnvironment}"; }
        }
        /// <summary>
        /// List of EntityMetadata objects for items being processed
        /// </summary>
        public List<EntityMetadata> Entities {
            get 
            {
                if (config.SelectedEntities?.Count > 0) 
                {
                    return AllEntities.Where(e => config.SelectedEntities.Contains(e.LogicalName)).ToList();
                }
                else {
                    return AllEntities; 
                }
            }
        }

        public List<EntityMetadata> AllEntities { get; set; }
        #endregion

        #region Helper methods 
        /// <summary>
        /// Helper for logging
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            var sb = new StringBuilder();

            string selectedEntitiesString = config.SelectedEntities == null ? "<UNDEFINED>" : string.Join(", ", config.SelectedEntities.ToArray());

            sb.AppendLine($"ActiveItemsOnly: {config.ActiveItemsOnly}")
                .AppendLine($"CreateFilter: {config.CreateFilter}")
                .AppendLine($"ModifyFilter: {config.ModifyFilter}")
                .AppendLine($"DateFilterOptions: {config.DateFilterOptions}")
                .AppendLine($"WebsiteFilter: {config.WebsiteFilter}")
                .AppendLine($"ImportFilename: {config.ImportFilename}")
                .AppendLine($"ExportFilename: {config.ExportFilename}")
                .AppendLine($"PriorDaysToRetrieve: {config.PriorDaysToRetrieve}")
                .AppendLine($"SourceUsername: {config.SourceUsername}")
                // .AppendLine($"SourcePassword: {config.SourcePassword}")
                .AppendLine($"TargetUsername: {config.TargetUsername}")
                // .AppendLine($"TargetPassword: {config.TargetPassword}")
                .AppendLine($"SourceEnvironment: {config.SourceEnvironment}")
                .AppendLine($"SourceConnectionString: {SourceConnectionString}")
                .AppendLine($"TargetEnvironment: {config.TargetEnvironment}")
                .AppendLine($"TargetConnectionString: {TargetConnectionString}")
                .AppendLine($"SelectedEntities:\n{selectedEntitiesString}");

            return sb.ToString();
        }
        #endregion

        #region Initialization

        /// <summary>
        /// Initialize a new instance using this "factory" method. Parses the command line args, loads settings file, sets some default values
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public static ExportSettings InitializeSettings(string[] args)
        {
            ExportSettings settings = new ExportSettings();

            var argsDict = new Dictionary<string, string>();

            // parse command line args. override any settings in the config file with the command line params 
            // first throw into dict so we can check for alternate Settings file
            foreach (var arg in args) {
                var argDelimiterPosition = arg.IndexOf(":");
                var argName = arg.Substring(1, argDelimiterPosition - 1);
                var argValue = arg.Substring(argDelimiterPosition + 1);
                argsDict.Add(argName.ToLower(), argValue);

                PortalMover.ReportProgress($"command line arg: Name: {argName.ToLower()}, Value: {argValue}");
            }

            // check to see if a new settings file was specified 
            if (argsDict.ContainsKey("settings")) {
                settings.SettingsFileName = argsDict["settings"];
            }

            // load the settings file first... then override values from the command line parameters
            if (File.Exists(settings.SettingsFileName)) {
                using (TextReader txtReader = new StreamReader(settings.SettingsFileName)) 
                {
                    var jsonString = txtReader.ReadToEnd();
                    settings.Config = MoverSettingsConfig.FromJson(jsonString);
                }
            }
            else {
                settings.Config = new MoverSettingsConfig();
                PortalMover.ReportProgress($"Unable to locate the configuration file {settings.SettingsFileName}.  Using Command Line args only.");
            }

            // update settings object with override values from command line
            // TODO error handling, throw exception on required elements, validation, etc
            foreach (var arg in argsDict)
            {
                switch (arg.Key)
                {
                    case "createdon":
                        if (DateTime.TryParse(arg.Value, out var created)) {
                            settings.Config.CreateFilter = created;
                        }
                        else {
                            settings.Config.CreateFilter = null;
                        }
                        break;
                    case "modifiedon":
                        if (DateTime.TryParse(arg.Value, out var mod)) {
                            settings.Config.ModifyFilter = mod;
                        }
                        else {
                            settings.Config.ModifyFilter = null;
                        }
                        break;
                    case "activeonly":
                        if (bool.TryParse(arg.Value, out var act)) {
                            settings.Config.ActiveItemsOnly = act;
                        }
                        break;
                    case "website":
                        if (!string.IsNullOrEmpty(arg.Value)) {
                            settings.Config.WebsiteFilter = Guid.Parse(arg.Value);
                        }
                        break;
                    case "exportfile":
                        settings.Config.ExportFilename = arg.Value;
                        break;
                    case "importfile":
                        settings.Config.ImportFilename = arg.Value;
                        break;
                    case "priordays":
                        if (int.TryParse(arg.Value, out var days)) {
                            settings.Config.PriorDaysToRetrieve = days;
                        }
                        else {
                            settings.Config.PriorDaysToRetrieve = null;
                        }
                        break;
                    // SOURCE
                    case "sourceenv":
                        settings.Config.SourceEnvironment = arg.Value;
                        break;
                    case "sourceuser":
                        settings.Config.SourceUsername = arg.Value;
                        break;
                    case "sourcepass":
                        settings.Config.SourcePassword = arg.Value;
                        break;

                    // TARGET
                    case "targetenv":
                        settings.Config.TargetEnvironment = arg.Value;
                        break;
                    case "targetuser":
                        settings.Config.TargetUsername = arg.Value;
                        break;
                    case "targetpass":
                        settings.Config.TargetPassword = arg.Value;
                        break;

                    case "datefilteroptions":
                        settings.Config.DateFilterOptions = DateFilterOptionsEnum.CreateAndModify;
                        if (Enum.TryParse<DateFilterOptionsEnum>(arg.Value, out DateFilterOptionsEnum df)) {
                            settings.Config.DateFilterOptions = df;
                        }
                        break;
                    case "cleanweb":
                        if (bool.TryParse(arg.Value, out var web)) {
                            settings.Config.CleanWebFiles = web;
                        }
                        break;
                    case "disableplugins":
                        if (bool.TryParse(arg.Value, out var plug)) {
                            settings.Config.DeactivateWebPagePlugins = plug;
                        }
                        break;
                    case "removejsrestriction":
                        if (bool.TryParse(arg.Value, out var js))
                        {
                            settings.Config.RemoveJavaScriptFileRestriction = js;
                        }
                        break;
                    case "removeformattedvalues":
                        if (bool.TryParse(arg.Value, out var rfv))
                        {
                            settings.Config.RemoveFormattedValues = rfv;
                        }
                        break;
                    case "exportinfolderstructure":
                        if (bool.TryParse(arg.Value, out var efs))
                        {
                            settings.Config.ExportInFolderStructure = efs;
                        }
                        break;
                    default:
                        break;
                }
            }

            if (settings.Config.DateFilterOptions == 0) {
                settings.Config.DateFilterOptions = DateFilterOptionsEnum.CreateAndModify;
            }

            // Only use the date part so we start at midnight on the day we want
            var now = DateTime.Now.Date;

            // if the prior days has been set, then override it with the calculated value 
            // this allows us to say, retrieve the last X days worth of records vs specifying a date
            if (settings.Config.PriorDaysToRetrieve != null) 
            {
                // override and update based on date filter options
                var dateFilter = now.AddDays((double)-settings.Config.PriorDaysToRetrieve);

                settings.Config.CreateFilter = dateFilter;
                settings.Config.ModifyFilter = dateFilter;
            }
            // null out dates we don't want 
            if (settings.Config.DateFilterOptions == DateFilterOptionsEnum.ModifyOnly) {
                settings.Config.CreateFilter = null;
            }

            if (settings.Config.DateFilterOptions == DateFilterOptionsEnum.CreateOnly) {
                settings.Config.ModifyFilter = null;
            }
            // make sure we have some required values set
            // TODO  look at impact of making this optional?
            if (settings.Config.WebsiteFilter == Guid.Empty) {
                throw new ArgumentNullException("WebsiteFilter must be specified");
            }

            if ((settings.Config.CreateFilter == null) && (settings.Config.ModifyFilter == null)) {
                throw new ArgumentNullException("Either CreateFilter or ModifyFilter must be specified");
            }

            // make sure we have some required values set
            if (string.IsNullOrEmpty(settings.Config.ImportFilename) && string.IsNullOrEmpty(settings.Config.ExportFilename)) {
                throw new ArgumentNullException("Either ImportFilename or ExportFilename must be specified");
            }

            if (!string.IsNullOrEmpty(settings.Config.ImportFilename) && string.IsNullOrEmpty(settings.Config.TargetEnvironment)) {
                throw new ArgumentNullException("TargetEnvironment must be specified if ImportFilename is specified");
            }

            if (!string.IsNullOrEmpty(settings.Config.ExportFilename) && string.IsNullOrEmpty(settings.Config.SourceEnvironment)) {
                throw new ArgumentNullException("SourceEnvironment must be specified if ExportFilename is specified");
            }

            if (!string.IsNullOrEmpty(settings.Config.ExportFilename))
            {
                if (string.IsNullOrEmpty(settings.Config.SourceUsername) || string.IsNullOrEmpty(settings.Config.SourcePassword))
                {
                    throw new ArgumentNullException("Source User name and password must be specified for Export");
                }
            }

            if (!string.IsNullOrEmpty(settings.Config.ImportFilename))
            {
                if (string.IsNullOrEmpty(settings.Config.TargetUsername) || string.IsNullOrEmpty(settings.Config.TargetPassword))
                {
                    throw new ArgumentNullException("Target User name and password must be specified for Import");
                }
            }
            // update the import/export file name with the date, if the mask is present
            if (!string.IsNullOrEmpty(settings.Config.ExportFilename)) {
                settings.Config.ExportFilename = string.Format(settings.Config.ExportFilename, now);
            }
            if (!string.IsNullOrEmpty(settings.Config.ImportFilename)) {
                settings.Config.ImportFilename = string.Format(settings.Config.ImportFilename, now);
            }

            // tell us what we have... 
            PortalMover.ReportProgress(settings.ToString());

            return settings;
        }
        #endregion
    }
}
