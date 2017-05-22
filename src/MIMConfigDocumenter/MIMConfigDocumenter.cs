//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="MIMConfigDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// MIM Configuration Documenter Utility
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace MIMConfigDocumenter
{
    using System;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Web.UI;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The MIMConfigDocumenter documents the configuration of an MIM Sync deployment.
    /// </summary>
    public class MIMConfigDocumenter : Documenter
    {
        /// <summary>
        /// The current pilot / test configuration directory.
        /// This is the revised / target configuration which has introduced new changes to the baseline / production environment. 
        /// </summary>
        private string pilotConfigDirectory;

        /// <summary>
        /// The current production configuration directory.
        /// The is the baseline / reference configuration on which the changes will be reported.
        /// </summary>
        private string productionConfigDirectory;

        /// <summary>
        /// The relative path of the current pilot / test configuration directory.
        /// </summary>
        private string pilotConfigRelativePath;

        /// <summary>
        /// The relative path of the current production configuration directory.
        /// </summary>
        private string productionConfigRelativePath;

        /// <summary>
        /// The configuration report file path
        /// </summary>
        private string configReportFilePath;

        /// <summary>
        /// Initializes a new instance of the <see cref="MIMConfigDocumenter"/> class.
        /// </summary>
        /// <param name="targetSystem">The target / pilot / test system.</param>
        /// <param name="referenceSystem">The reference / baseline / production system.</param>
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "GlobalSettings", Justification = "Reviewed.")]
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "SynchronizationRules", Justification = "Reviewed.")]
        public MIMConfigDocumenter(string targetSystem, string referenceSystem)
        {
            Logger.Instance.WriteMethodEntry("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", targetSystem, referenceSystem);

            try
            {
                this.pilotConfigRelativePath = targetSystem;
                this.productionConfigRelativePath = referenceSystem;
                this.ReportFileName = Documenter.GetTempFilePath("Report.tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath("Report.TOC.tmp.html");

                var rootDirectory = Directory.GetCurrentDirectory().TrimEnd('\\');

                this.pilotConfigDirectory = string.Format(CultureInfo.InvariantCulture, @"{0}\Data\{1}", rootDirectory, this.pilotConfigRelativePath);
                this.productionConfigDirectory = string.Format(CultureInfo.InvariantCulture, @"{0}\Data\{1}", rootDirectory, this.productionConfigRelativePath);
                this.configReportFilePath = Documenter.ReportFolder + @"\" + Documenter.GetReportFileBaseName(this.pilotConfigRelativePath, this.productionConfigRelativePath) + "_Consolidated_report.html";
            }
            finally
            {
                Logger.Instance.WriteMethodExit("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", targetSystem, referenceSystem);
            }
        }

        /// <summary>
        /// Generates the report.
        /// </summary>
        public void GenerateReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var report = this.GetReport();

                Logger.Instance.WriteInfo("Writing Consolidated Report...");

                this.WriteReport("FIM/MIM Configuration", report.Item1, report.Item2, this.pilotConfigRelativePath, this.productionConfigRelativePath, this.configReportFilePath);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the MIM Sync configuration report.
        /// </summary>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            Tuple<string, string> report;

            try
            {
                Tuple<string, string> syncReportTuple = new Tuple<string, string>(string.Empty, string.Empty);
                Tuple<string, string> serviceDocumenterTuple = new Tuple<string, string>(string.Empty, string.Empty);

                try
                {
                    var syncDocumenter = new MIMSyncConfigDocumenter(this.pilotConfigRelativePath, this.productionConfigRelativePath);
                    syncReportTuple = syncDocumenter.GetReport();
                }
                catch (FileNotFoundException e)
                {
                    Logger.Instance.WriteError(e.ToString());
                }

                try
                {
                    var serviceDocumenter = new MIMServiceConfigDocumenter(this.pilotConfigRelativePath, this.productionConfigRelativePath);
                    serviceDocumenterTuple = serviceDocumenter.GetReport();
                }
                catch (FileNotFoundException e)
                {
                    Logger.Instance.WriteError(e.ToString());
                }

                report = new Tuple<string, string>(syncReportTuple.Item1 + serviceDocumenterTuple.Item1, syncReportTuple.Item2 + serviceDocumenterTuple.Item2);
            }
            catch (Exception e)
            {
                throw Logger.Instance.ReportError(e);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }

            return report;
        }
    }
}
