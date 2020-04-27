using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using Microsoft.Xrm.Sdk.Metadata;
using System.Text;
using System.IO;

using Newtonsoft.Json;
using Newtonsoft.Json.Converters;


namespace PortalRecordsMover.AppCode
{
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
            get { return $"RequireNewInstance=True;AuthType=Office365;Username={config.Username}; Password={config.Password};Url={config.SourceEnvironment}"; }
        }
        public string TargetConnectionString {
            get { return $"RequireNewInstance=True;AuthType=Office365;Username={config.Username}; Password={config.Password};Url={config.TargetEnvironment}"; }
        }

        public List<EntityMetadata> Entities {
            get { return AllEntities.Where(e => config.SelectedEntities.Contains(e.LogicalName)).ToList(); }
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

            sb.AppendLine($"ActiveItemsOnly: {config.ActiveItemsOnly}")
                .AppendLine($"CreateFilter: {config.CreateFilter}")
                .AppendLine($"ModifyFilter: {config.ModifyFilter}")
                .AppendLine($"DateFilterOptions: {config.DateFilterOptions}")
                .AppendLine($"BatchCount: {config.BatchCount}")
                .AppendLine($"WebsiteFilter: {config.WebsiteFilter}")
                .AppendLine($"ImportFilename: {config.ImportFilename}")
                .AppendLine($"ExportFilename: {config.ExportFilename}")
                .AppendLine($"PriorDaysToRetrieve: {config.PriorDaysToRetrieve}")
                .AppendLine($"Username: {config.Username}")
                .AppendLine($"Password: ********")
                .AppendLine($"SourceEnvironment: {config.SourceEnvironment}")
                .AppendLine($"SourceConnectionString: {SourceConnectionString}")
                .AppendLine($"TargetEnvironment: {config.TargetEnvironment}")
                .AppendLine($"TargetConnectionString: {TargetConnectionString}")
                .AppendLine($"SelectedEntities:\n{string.Join(", ", config.SelectedEntities.ToArray())}");

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

            // load the settings file
            if (File.Exists(settings.SettingsFileName)) {
                using (TextReader txtReader = new StreamReader(settings.SettingsFileName)) 
                {
                    var jsonString = txtReader.ReadToEnd();
                    settings.Config = MoverSettingsConfig.FromJson(jsonString);
                }
            }
            else {
                PortalMover.ReportProgress($"Unable to locate the configruation file {settings.SettingsFileName}.  Using Command Line args only.");
            }

            // update settings object with override values from command line
            // TODO error handling, throw exception on required elements, validation, etc
            foreach (var arg in argsDict)
            {
                DateTime dt;
                switch (arg.Key)
                {
                    case "createdon":
                        if (DateTime.TryParse(arg.Value, out dt)) {
                            settings.Config.CreateFilter = dt;
                        }
                        else {
                            settings.Config.CreateFilter = null;
                        }
                        break;
                    case "modifiedon":
                        if (DateTime.TryParse(arg.Value, out dt)) {
                            settings.Config.ModifyFilter = dt;
                        }
                        else {
                            settings.Config.ModifyFilter = null;
                        }
                        break;
                    case "activeonly":
                        if (bool.TryParse(arg.Value, out var b)) {
                            settings.Config.ActiveItemsOnly = b;
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
                        if (int.TryParse(arg.Value, out var i)) {
                            settings.Config.PriorDaysToRetrieve = i;
                        }
                        else {
                            settings.Config.PriorDaysToRetrieve = null;
                        }
                        break;
                    case "targetenv":
                        settings.Config.TargetEnvironment = arg.Value;
                        break;
                    case "sourceenv":
                        settings.Config.SourceEnvironment = arg.Value;
                        break;
                    case "user":
                        settings.Config.Username = arg.Value;
                        break;
                    case "pass":
                        settings.Config.Password = arg.Value;
                        break;

                    case "batchcount":
                        settings.Config.BatchCount = 10;
                        if (int.TryParse(arg.Value, out var bc)) {
                            settings.Config.BatchCount = bc;
                        }
                        break;

                    case "datefilteroptions":
                        settings.Config.DateFilterOptions = DateFilterOptionsEnum.CreateAndModify;
                        if (Enum.TryParse<DateFilterOptionsEnum>(arg.Value, out DateFilterOptionsEnum df)) {
                            settings.Config.DateFilterOptions = df;
                        }
                        break;

                    default:
                        break;
                }
            }

            // defaults~
            if (settings.Config.BatchCount == 0) {
                settings.Config.BatchCount = 10;
            }
            if (settings.Config.DateFilterOptions == 0) {
                settings.Config.DateFilterOptions = DateFilterOptionsEnum.CreateAndModify;
            }

            // Only use the date part so we start at midnight on the day we want
            var now = DateTime.Now.Date;

            // if the prior days has been set, then override it with the calculated value 
            // this allows us to say, retrieve the last X days worth of records vs specifying a date
            if (settings.Config.PriorDaysToRetrieve != null || settings.Config.PriorDaysToRetrieve > 0) 
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

            if (string.IsNullOrEmpty(settings.Config.Username) || string.IsNullOrEmpty(settings.Config.Password)) {
                throw new ArgumentNullException("User name and password must be specified");
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
