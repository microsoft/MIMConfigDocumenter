//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="FIMConnectorDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// MIM FIM Connector Configuration Documenter
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace MIMConfigDocumenter
{
    using System;
    using System.Collections.Specialized;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.Linq;
    using System.Web.UI;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The FIMConnectorDocumenter documents the configuration of a FIM connector.
    /// </summary>
    internal class FIMConnectorDocumenter : ConnectorDocumenter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FIMConnectorDocumenter"/> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="connectorName">The connector name.</param>
        /// <param name="configEnvironment">The environment in which the config element exists.</param>
        public FIMConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, ConfigEnvironment configEnvironment)
            : base(pilotXml, productionXml, connectorName, configEnvironment)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportFileName = Documenter.GetTempFilePath(this.ConnectorName + ".tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath(this.ConnectorName + ".TOC.tmp.html");
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the FIM connector configuration report.
        /// </summary>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.WriteConnectorReportHeader();

                this.ProcessConnectorProperties();

                this.ProcessFIMConnectivityInformation();

                this.ProcessConnectorSelectedObjectTypes();
                this.ProcessConnectorSelectedAttributes();
                this.ProcessConnectorFilterRules();
                this.ProcessFIMObjectTypeMappings();
                this.ProcessConnectorAttributeFlows();
                this.ProcessConnectorDeprovisioningConfiguration();
                this.ProcessConnectorRunProfiles();

                return this.GetReportTuple();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();

                // Clear Logger call context items
                Logger.ClearContextItem(ConnectorDocumenter.LoggerContextItemConnectorName);
                Logger.ClearContextItem(ConnectorDocumenter.LoggerContextItemConnectorGuid);
                Logger.ClearContextItem(ConnectorDocumenter.LoggerContextItemConnectorCategory);
            }
        }

        #region Connectivity Information

        /// <summary>
        /// Processes the FIM connectivity information.
        /// </summary>
        protected void ProcessFIMConnectivityInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connectivity Information");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1= Setting Display Order, 2 = Setting, 3 = Configuration

                this.FillFIMConnectivityInformationDataSet(true);
                this.FillFIMConnectivityInformationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintFIMConnectivityInformation();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the FIM connectivity information data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillFIMConnectivityInformationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var table = dataSet.Tables[0];

                    var server = (string)connector.XPathSelectElement("private-configuration/fimma-configuration/connection-info/server");
                    var database = (string)connector.XPathSelectElement("private-configuration/fimma-configuration/connection-info/databasename");
                    var serviceHost = (string)connector.XPathSelectElement("private-configuration/fimma-configuration/connection-info/serviceHost");
                    var integratedAuth = ((string)connector.XPathSelectElement("private-configuration/fimma-configuration/connection-info/authentication") ?? string.Empty).Equals("integrated", StringComparison.OrdinalIgnoreCase);
                    var userName = (string)connector.XPathSelectElement("private-configuration/fimma-configuration/connection-info/user");
                    var domain = (string)connector.XPathSelectElement("private-configuration/fimma-configuration/connection-info/domain");

                    Documenter.AddRow(table, new object[] { 1, "Server", server });
                    Documenter.AddRow(table, new object[] { 2, "Database", database });
                    Documenter.AddRow(table, new object[] { 3, "FIM Service base address", serviceHost });
                    Documenter.AddRow(table, new object[] { 4, "Authentication Mode", integratedAuth ? "Windows integrated authentication" : "SQL authentication" });
                    Documenter.AddRow(table, new object[] { 5, "User name", userName });
                    Documenter.AddRow(table, new object[] { 6, "Password", "*****" });
                    Documenter.AddRow(table, new object[] { 7, "Domain", domain });

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the FIM connectivity information.
        /// </summary>
        protected void PrintFIMConnectivityInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Connectivity Information";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Setting", 50 }, { "Configuration", 50 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Connectivity Information

        #region Object Type Mappings

        /// <summary>
        /// Processes the FIM object type mappings information.
        /// </summary>
        protected void ProcessFIMObjectTypeMappings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connectivity Information");

                this.CreateSimpleSettingsDataSets(2); // 1 = Setting, 2 = Configuration

                this.FillFIMObjectTypeMappingsDataSet(true);
                this.FillFIMObjectTypeMappingsDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintFIMObjectTypeMappings();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the FIM object type mapping data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillFIMObjectTypeMappingsDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var table = dataSet.Tables[0];

                    var projectionRules = from projectionRule in connector.XPathSelectElements("projection/class-mapping")
                                          let sourceObjectType = (string)projectionRule.Attribute("cd-object-type")
                                          orderby sourceObjectType
                                          select projectionRule;

                    if (projectionRules.Count() == 0)
                    {
                        return;
                    }

                    foreach (var projectionRule in projectionRules)
                    {
                        var sourceObjectType = (string)projectionRule.Attribute("cd-object-type");
                        var metaverseObjectType = (string)projectionRule.Element("mv-object-type") ?? string.Empty;

                        Documenter.AddRow(table, new object[] { sourceObjectType, metaverseObjectType });
                    }

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the FIM object type mappings.
        /// </summary>
        protected void PrintFIMObjectTypeMappings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Object Type Mappings";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Data Source Object Type", 50 }, { "Metaverse Object Type", 50 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Object Type Mappings

        #region Run Profiles

        /// <summary>
        /// Fills the FIM run profile data set.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected override void FillConnectorRunProfileDataSet(string runProfileName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Run Profile Name: '{0}'. Pilot Config: '{1}'.", runProfileName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    base.FillConnectorRunProfileDataSet(runProfileName, pilotConfig);

                    var table = dataSet.Tables[0];
                    var table2 = dataSet.Tables[1];

                    var runProfileSteps = connector.XPathSelectElements("ma-run-data/run-configuration[name = '" + runProfileName + "']/configuration/step");

                    for (var stepIndex = 1; stepIndex <= runProfileSteps.Count(); ++stepIndex)
                    {
                        var runProfileStep = runProfileSteps.ElementAt(stepIndex - 1);

                        var timeout = (string)runProfileStep.XPathSelectElement("custom-data/fimma-step-data/time-limit");
                        if (!string.IsNullOrEmpty(timeout))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Timeout (seconds)", timeout, 1000 });
                        }
                    }

                    table.AcceptChanges();
                    table2.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Run Profile Name: '{0}'. Pilot Config: '{1}'.", runProfileName, pilotConfig);
            }
        }

        #endregion Run Profiles
    }
}
