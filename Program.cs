using System;
using System.IO;
using System.Text;

using MscrmTools.PortalRecordsMover.AppCode;
using System.Threading.Tasks;

// Port of the Portal Records Mover http://XrmToolbox.com plugin by Tanguy Touzard
// Original project: PortalRecordsMover: https://github.com/MscrmTools/MscrmTools.PortalRecordsMover/  
// This project is console application version of the MscrmTools.PortalRecordsMover plugin included with the XrmToolbox.

namespace MscrmTools.PortalRecordsMover
{
    public class PortalMover
    {
        public const string DT_FORMAT = "yyyy-MM-dd HH.mm.ss";

        static StringBuilder _logger = new StringBuilder();

        public static StringBuilder Logger { get => _logger; }

        static void Main(string[] args) {
            // process the settings from file and then override with command line args, if any
            ReportProgress("Initializing the ExportSettings");
            ExportSettings settings = ExportSettings.InitializeSettings(args);

            if (!string.IsNullOrEmpty(settings.ExportFilename)) {
                ReportProgress($"Beginning the export - SourceEnvironment: {settings.SourceEnvironment}, ExportFilename:{settings.ExportFilename}");
                var exporter = new Exporter(settings);
                exporter.Export();
            }

            // if ImportFile is not null, then perform the import
            if (!string.IsNullOrEmpty(settings.ImportFilename)) {
                ReportProgress($"Beginning the Import - TargetEnvironment: {settings.TargetEnvironment}, ImportFilename:{settings.ImportFilename}");
                var importer = new Importer(settings);
                importer.Import();
            }

            // output the log file
            SaveLogFile().Wait();
            Console.ReadKey();
        }

        /// <summary>
        /// Save the log!
        /// </summary>
        /// <returns></returns>
        static async Task SaveLogFile()
        {
            var now = DateTime.Now.ToString(DT_FORMAT);
            var logFile = $"PortalRecordsMoverApp {now}.log";

            byte[] encodedText = Encoding.Unicode.GetBytes(Logger.ToString());

            using (FileStream sourceStream = new FileStream(logFile,FileMode.Append, FileAccess.Write, FileShare.None, bufferSize: 4096, useAsync: true)) {
                await sourceStream.WriteAsync(encodedText, 0, encodedText.Length);
            };
            ReportProgress($"Logfile saved to {logFile}");
        }

        /// <summary>
        /// Helper method to track logging for all of the processes 
        /// </summary>
        /// <param name="message"></param>
        public static void ReportProgress(string message)
        {
            var now = DateTime.Now.ToString(DT_FORMAT);
            message = $"{now}: {message}";
            Logger.AppendLine(message);
            Console.WriteLine(message);
        }
    }
}
