using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using Microsoft.Xrm.Sdk.Metadata;
using System.Text;
using System.IO;

namespace MscrmTools.PortalRecordsMover.AppCode
{
    public class ExportSettings
    {
        public ExportSettings()
        {
            SelectedEntities = new List<string>();
            SettingsFileName = "PortalRecordsMover.settings.xml";
        }

        #region Properties 
        public enum DateFilterOptionsEnum {
            CreateOnly = 1,
            ModifyOnly, 
            CreateAndModify
        }

        public string SettingsFileName { get; set; }

        // Use this number of days to calculate CreateFilter and ModifyFilter
        public int? PriorDaysToRetrieve { get; set; }

        public bool ActiveItemsOnly { get; set; }
        public DateTime? CreateFilter { get; set; }
        public DateTime? ModifyFilter { get; set; }
        public int BatchCount { get; set; }
        public DateFilterOptionsEnum DateFilterOptions { get; set; }
        public Guid WebsiteFilter { get; set; }
        public string ExportFilename { get; set; }
        public string ImportFilename { get; set; }
        public string SourceEnvironment { get; set; }
        public string TargetEnvironment { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }

        public List<string> SelectedEntities { get; set; }

        public List<WebsiteIdMap> WebsiteIdMaping { get; set; }

        [XmlIgnore]
        public string SourceConnectionString {
            get { return $"RequireNewInstance=True;AuthType=Office365;Username={Username}; Password={Password};Url={SourceEnvironment}"; }
        }
        [XmlIgnore]
        public string TargetConnectionString {
            get { return $"RequireNewInstance=True;AuthType=Office365;Username={Username}; Password={Password};Url={TargetEnvironment}"; }
        }

        [XmlIgnore]
        public List<EntityMetadata> Entities {
            get { return AllEntities.Where(e => SelectedEntities.Contains(e.LogicalName)).ToList(); }
        }

        [XmlIgnore]
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

            sb.AppendLine($"ActiveItemsOnly: {ActiveItemsOnly}")
                .AppendLine($"CreateFilter: {CreateFilter}")
                .AppendLine($"ModifyFilter: {ModifyFilter}")
                .AppendLine($"DateFilterOptions: {DateFilterOptions}")
                .AppendLine($"BatchCount: {BatchCount}")
                .AppendLine($"WebsiteFilter: {WebsiteFilter}")
                .AppendLine($"ImportFilename: {ImportFilename}")
                .AppendLine($"ExportFilename: {ExportFilename}")
                .AppendLine($"PriorDaysToRetrieve: {PriorDaysToRetrieve}")
                .AppendLine($"Username: {Username}")
                .AppendLine($"Password: {Password}")
                .AppendLine($"SourceEnvironment: {SourceEnvironment}")
                .AppendLine($"SourceConnectionString: {SourceConnectionString}")
                .AppendLine($"TargetEnvironment: {TargetEnvironment}")
                .AppendLine($"TargetConnectionString: {TargetConnectionString}")
                .AppendLine($"SelectedEntities:\n{string.Join(", ", SelectedEntities.ToArray())}");

            //foreach (var entity in SelectedEntities)
            //{
            //    sb.AppendLine($"\t{entity}");
            //}

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
                using (TextReader txtReader = new StreamReader(settings.SettingsFileName)) {
                    XmlSerializer xmlSerializer = new XmlSerializer(typeof(ExportSettings));
                    settings = (ExportSettings)xmlSerializer.Deserialize(txtReader);
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
                            settings.CreateFilter = dt;
                        }
                        else {
                            settings.CreateFilter = null;
                        }
                        break;
                    case "modifiedon":
                        if (DateTime.TryParse(arg.Value, out dt)) {
                            settings.ModifyFilter = dt;
                        }
                        else {
                            settings.ModifyFilter = null;
                        }
                        break;
                    case "activeonly":
                        if (bool.TryParse(arg.Value, out var b)) {
                            settings.ActiveItemsOnly = b;
                        }
                        break;
                    case "website":
                        if (!string.IsNullOrEmpty(arg.Value)) {
                            settings.WebsiteFilter = Guid.Parse(arg.Value);
                        }
                        break;
                    case "exportfile":
                            settings.ExportFilename = arg.Value;
                        break;
                    case "importfile":
                            settings.ImportFilename = arg.Value;
                        break;
                    case "priordays":
                        if (int.TryParse(arg.Value, out var i)) {
                            settings.PriorDaysToRetrieve = i;
                        }
                        else {
                            settings.PriorDaysToRetrieve = null;
                        }
                        break;
                    case "targetenv":
                            settings.TargetEnvironment = arg.Value;
                        break;
                    case "sourceenv":
                            settings.SourceEnvironment = arg.Value;
                        break;
                    case "user":
                            settings.SourceEnvironment = arg.Value;
                        break;
                    case "pass":
                        settings.SourceEnvironment = arg.Value;
                        break;

                    case "batchcount":
                        settings.BatchCount = 10;
                        if (int.TryParse(arg.Value, out var bc)) {
                            settings.BatchCount = bc;
                        }
                        break;

                    case "datefilteroptions":
                        settings.DateFilterOptions = DateFilterOptionsEnum.CreateAndModify;
                        if (Enum.TryParse<DateFilterOptionsEnum>(arg.Value, out DateFilterOptionsEnum df)) {
                            settings.DateFilterOptions = df;
                        }
                        break;

                    default:
                        break;
                }
            }

            // defaults~
            if (settings.BatchCount == 0) {
                settings.BatchCount = 10;
            }
            if (settings.DateFilterOptions == 0) {
                settings.DateFilterOptions = DateFilterOptionsEnum.CreateAndModify;
            }

            // Only use the date part so we start at midnight on the day we want
            var now = DateTime.Now.Date;

            // if the prior days has been set, then override it with the calculated value 
            // this allows us to say, retrieve the last X days worth of records vs specifying a date
            if (settings.PriorDaysToRetrieve != null) 
            {
                // override and update based on date filter options
                var dateFilter = now.AddDays((double)-settings.PriorDaysToRetrieve);

                settings.CreateFilter = dateFilter;
                settings.ModifyFilter = dateFilter;
            }
            // null out dates we don't want 
            if (settings.DateFilterOptions == DateFilterOptionsEnum.ModifyOnly) {
                settings.CreateFilter = null;
            }

            if (settings.DateFilterOptions == DateFilterOptionsEnum.CreateOnly) {
                settings.ModifyFilter = null;
            }
            // make sure we have some required values set
            // TODO  look at impact of making this optional?
            if (settings.WebsiteFilter == Guid.Empty) {
                throw new ArgumentNullException("WebsiteFilter must be specified");
            }

            if ((settings.CreateFilter == null) && (settings.ModifyFilter == null)) {
                throw new ArgumentNullException("Either CreateFilter or ModifyFilter must be specified");
            }

            // make sure we have some required values set
            if (string.IsNullOrEmpty(settings.ImportFilename) && string.IsNullOrEmpty(settings.ExportFilename)) {
                throw new ArgumentNullException("Either ImportFilename or ExportFilename must be specified");
            }

            if (!string.IsNullOrEmpty(settings.ImportFilename) && string.IsNullOrEmpty(settings.TargetEnvironment)) {
                throw new ArgumentNullException("TargetEnvironment must be specified if ImportFilename is specified");
            }

            if (!string.IsNullOrEmpty(settings.ExportFilename) && string.IsNullOrEmpty(settings.SourceEnvironment)) {
                throw new ArgumentNullException("SourceEnvironment must be specified if ExportFilename is specified");
            }

            if (string.IsNullOrEmpty(settings.Username) || string.IsNullOrEmpty(settings.Password)) {
                throw new ArgumentNullException("User name and password must be specified");
            }

            // update the import/export file name with the date, if the mask is present
            if (!string.IsNullOrEmpty(settings.ExportFilename)) {
                settings.ExportFilename = string.Format(settings.ExportFilename, now);
            }
            if (!string.IsNullOrEmpty(settings.ImportFilename)) {
                settings.ImportFilename = string.Format(settings.ImportFilename, now);
            }

            // tell us what we have... 
            PortalMover.ReportProgress(settings.ToString());

            return settings;
        }
        #endregion
    }

    /// <summary>
    /// Helper class to capture the mapping of website IDs from one environment to the other
    /// </summary>
    public class WebsiteIdMap
    {
        public Guid SourceId { get; set; }
        public Guid TargetId { get; set; }
    }
}
