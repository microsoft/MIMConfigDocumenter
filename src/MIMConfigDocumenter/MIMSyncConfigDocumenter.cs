//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="MIMSyncConfigDocumenter.cs" company="Microsoft">
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
    /// The MIMSyncConfigDocumenter documents the configuration of an MIM Sync deployment.
    /// </summary>
    public class MIMSyncConfigDocumenter : Documenter
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
        /// Initializes a new instance of the <see cref="MIMSyncConfigDocumenter"/> class.
        /// </summary>
        /// <param name="targetSystem">The target / pilot / test system.</param>
        /// <param name="referenceSystem">The reference / baseline / production system.</param>
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "GlobalSettings", Justification = "Reviewed.")]
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "SynchronizationRules", Justification = "Reviewed.")]
        public MIMSyncConfigDocumenter(string targetSystem, string referenceSystem)
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
                this.configReportFilePath = Documenter.ReportFolder + @"\" + Documenter.GetReportFileBaseName(this.pilotConfigRelativePath, this.productionConfigRelativePath) + "_Sync_report.html";

                this.ValidateInput();
                this.MergeSyncExports();
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
                this.WriteReport("Synchronisation Service Configuration", report.Item1, report.Item2, this.pilotConfigRelativePath, this.productionConfigRelativePath, this.configReportFilePath);
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
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

                var sectionTitle = "Synchronization Service Configuration";
                this.WriteSectionHeader(sectionTitle, 1);

                this.ProcessMetaverseConfiguration();
                this.ProcessConnectorConfigurations();
            }
            catch (Exception e)
            {
                throw Logger.Instance.ReportError(e);
            }
            finally
            {
                report = this.GetReportTuple();

                Logger.Instance.WriteMethodExit();
            }

            return report;
        }

        /// <summary>
        /// Validates the input.
        /// </summary>
        private void ValidateInput()
        {
            Logger.Instance.WriteMethodEntry("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", this.pilotConfigRelativePath, this.productionConfigRelativePath);

            try
            {
                if (!Directory.Exists(this.pilotConfigDirectory))
                {
                    var error = string.Format(CultureInfo.CurrentUICulture, DocumenterResources.PilotConfigurationDirectoryNotFound, this.pilotConfigDirectory);
                    throw Logger.Instance.ReportError(new FileNotFoundException(error));
                }
                else if (!Directory.Exists(this.pilotConfigDirectory + @"\" + "SyncConfig"))
                {
                    var error = string.Format(CultureInfo.CurrentUICulture, DocumenterResources.SyncConfigurationDirectoryNotFound, this.pilotConfigDirectory);
                    throw Logger.Instance.ReportError(new FileNotFoundException(error));
                }

                if (!Directory.Exists(this.productionConfigDirectory))
                {
                    var error = string.Format(CultureInfo.CurrentUICulture, DocumenterResources.ProductionConfigurationDirectoryNotFound, this.productionConfigDirectory);
                    throw Logger.Instance.ReportError(new FileNotFoundException(error));
                }
                else if (!Directory.Exists(this.productionConfigDirectory + @"\" + "SyncConfig"))
                {
                    var error = string.Format(CultureInfo.CurrentUICulture, DocumenterResources.SyncConfigurationDirectoryNotFound, this.productionConfigDirectory);
                    throw Logger.Instance.ReportError(new FileNotFoundException(error));
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("TargetSystem: '{0}'. ReferenceSystem: '{1}'.", this.pilotConfigRelativePath, this.productionConfigRelativePath);
            }
        }

        /// <summary>
        /// Merges the ADSync configuration export XML files into a single XML file.
        /// </summary>
        private void MergeSyncExports()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PilotXml = this.MergeSyncConfigurationExports(true);
                this.ProductionXml = this.MergeSyncConfigurationExports(false);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Merges the Sync configuration exports.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, indicates that this is a pilot configuration. Otherwise, this is a production configuration.</param>
        /// <returns>
        /// An <see cref="XElement" /> object representing the combined configuration XML object.
        /// </returns>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Xml.Linq.XElement.Parse(System.String)", Justification = "Template XML is not localizable.")]
        private XElement MergeSyncConfigurationExports(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var configDirectory = pilotConfig ? this.pilotConfigDirectory : this.productionConfigDirectory;
                var templateXml = string.Format(CultureInfo.InvariantCulture, "<Root><{0}><SyncConfig/></{0}></Root>", pilotConfig ? "Pilot" : "Production");
                var configXml = XElement.Parse(templateXml);

                var syncConfig = configXml.XPathSelectElement("*//SyncConfig");
                foreach (var file in Directory.EnumerateFiles(configDirectory + "/SyncConfig", "*.xml"))
                {
                    syncConfig.Add(XElement.Load(file));
                }

                return configXml;
            }
            finally
            {
                Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Processes the metaverse configuration.
        /// </summary>
        private void ProcessMetaverseConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var metaverseDocumenter = new MetaverseDocumenter(this.PilotXml, this.ProductionXml);
                var report = metaverseDocumenter.GetReport();

                this.ReportWriter.Write(report.Item1);
                this.ReportToCWriter.Write(report.Item2);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the connector configurations.
        /// </summary>
        private void ProcessConnectorConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                const string XPath = "//ma-data";

                var pilot = this.PilotXml.XPathSelectElements(XPath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(XPath, Documenter.NamespaceManager);

                // Sort by name
                pilot = from connector in pilot
                        let name = (string)connector.Element("name")
                        orderby name
                        select connector;

                // Sort by name
                production = from connector in production
                             let name = (string)connector.Element("name")
                             orderby name
                             select connector;

                var pilotConnectors = from pilotConnector in pilot
                                      select (string)pilotConnector.Element("name");

                foreach (var connector in pilot)
                {
                    var configEnvironment = production.Any(productionConnector => (string)productionConnector.Element("name") == (string)connector.Element("name")) ? ConfigEnvironment.PilotAndProduction : ConfigEnvironment.PilotOnly;
                    this.ProcessConnectorConfiguration(connector, configEnvironment);
                }

                production = production.Where(productionConnector => !pilotConnectors.Contains((string)productionConnector.Element("name")));

                foreach (var connector in production)
                {
                    this.ProcessConnectorConfiguration(connector, ConfigEnvironment.ProductionOnly);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the connector configuration of specified connector.
        /// </summary>
        /// <param name="connector"> The connector config node</param>
        /// <param name="configEnvironment">The config environment.</param>
        private void ProcessConnectorConfiguration(XElement connector, ConfigEnvironment configEnvironment)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var connectorName = (string)connector.Element("name");
                var connectorCategory = (string)connector.Element("category");
                ConnectorDocumenter connectorDocumenter;

                switch (connectorCategory.ToUpperInvariant())
                {
                    case "AD":
                        connectorDocumenter = new ActiveDirectoryConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                        break;
                    case "AD GAL":
                        connectorDocumenter = new ActiveDirectoryGALConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                        break;
                    case "ADAM":
                        connectorDocumenter = new ActiveDirectoryLDSConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                        break;
                    case "FIM":
                        connectorDocumenter = new FIMConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                        break;
                    case "MSSQL":
                    case "ORACLE":
                    case "DB2":
                        connectorDocumenter = new DatabaseConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                        break;
                    case "EXTENSIBLE2":
                        {
                            var connectorSubType = (string)connector.Element("subtype");
                            switch (connectorSubType.ToUpperInvariant())
                            {
                                case "WINDOWS AZURE ACTIVE DIRECTORY (MICROSOFT)":
                                    connectorDocumenter = new AzureActiveDirectoryConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                case "POWERSHELL (MICROSOFT)":
                                    connectorDocumenter = new PowerShellConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                case "GENERIC SQL (MICROSOFT)":
                                    connectorDocumenter = new GenericSqlConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                case "GENERIC LDAP (MICROSOFT)":
                                    connectorDocumenter = new GenericLdapConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                case "MICROSOFT WEB SERVICE ECMA2 CONNECTOR (MICROSOFT)":
                                    connectorDocumenter = new WebServicesConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                                default:
                                    if (!string.IsNullOrEmpty(connectorSubType))
                                    {
                                        Logger.Instance.WriteWarning("ECMA2 Connector of subtype '{0}' is currently not supported. The connector '{1}' with be treated as a generic ECMA2 connector.", connectorSubType, connectorName);
                                    }

                                    connectorDocumenter = new Extensible2ConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                                    break;
                            }
                        }

                        break;
                    default:
                        Logger.Instance.WriteWarning("Connector of type '{0}' is currently not supported. The connector '{1}' with be treated as a generic ECMA2 connector. Documentation may not be complete.", connectorCategory, connectorName);
                        connectorDocumenter = new Extensible2ConnectorDocumenter(this.PilotXml, this.ProductionXml, connectorName, configEnvironment);
                        break;
                }

                var report = connectorDocumenter.GetReport();
                this.ReportWriter.Write(report.Item1);
                this.ReportToCWriter.Write(report.Item2);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }
    }
}
