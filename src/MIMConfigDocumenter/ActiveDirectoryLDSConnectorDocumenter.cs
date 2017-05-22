//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ActiveDirectoryLDSConnectorDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// MIM Configuration Documenter
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
    /// The ActiveDirectoryLDSConnectorDocumenter documents the configuration of an Active Directory LDS connector.
    /// </summary>
    internal class ActiveDirectoryLDSConnectorDocumenter : ActiveDirectoryConnectorDocumenter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ActiveDirectoryLDSConnectorDocumenter"/> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="connectorName">The name.</param>
        /// <param name="configEnvironment">The environment in which the config element exists.</param>
        public ActiveDirectoryLDSConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, ConfigEnvironment configEnvironment)
            : base(pilotXml, productionXml, connectorName, configEnvironment)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the active directory connector configuration report.
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

                this.ProcessActiveDirectoryLDSConnectionInformation();
                this.ProcessActiveDirectoryPartitions();
                this.ProcessConnectorProvisioningHierarchyConfiguration();
                this.ProcessConnectorSelectedObjectTypes();
                this.ProcessConnectorSelectedAttributes();
                this.ProcessConnectorFilterRules();
                this.ProcessConnectorJoinAndProjectionRules();
                this.ProcessConnectorAttributeFlows();
                this.ProcessConnectorDeprovisioningConfiguration();
                this.ProcessActiveDirectoryExtensionsConfiguration();
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

        #region AD LDS Connection Information

        /// <summary>
        /// Processes the active directory connection information.
        /// </summary>
        protected void ProcessActiveDirectoryLDSConnectionInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Active Directory LDS Connection Information.");

                // Forest Connection Information
                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Setting, 3 = Configuration

                this.FillActiveDirectoryLDSConnectionInformationDataSet(true);
                this.FillActiveDirectoryLDSConnectionInformationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintActiveDirectoryLDSConnectionInformation();

                // Forest Connection Option
                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Setting, 3 = Configuration

                this.FillActiveDirectoryConnectionOptionDataSet(true);
                this.FillActiveDirectoryConnectionOptionDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintActiveDirectoryConnectionOption();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region AD Forest Information

        /// <summary>
        /// Fills the active directory LDS connection information data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillActiveDirectoryLDSConnectionInformationDataSet(bool pilotConfig)
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

                    var forestName = (string)connector.XPathSelectElement("private-configuration/adma-configuration/forest-name");
                    var port = (string)connector.XPathSelectElement("private-configuration/adma-configuration/forest-port");
                    var userName = (string)connector.XPathSelectElement("private-configuration/adma-configuration/forest-login-user");
                    var userDomain = (string)connector.XPathSelectElement("private-configuration/adma-configuration/forest-login-domain");

                    Documenter.AddRow(table, new object[] { 1, "Server Name", forestName });
                    Documenter.AddRow(table, new object[] { 2, "Port", port });
                    Documenter.AddRow(table, new object[] { 3, "User Name", userName });
                    Documenter.AddRow(table, new object[] { 4, "Domain", userDomain });

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the active directory LDS connection information.
        /// </summary>
        protected void PrintActiveDirectoryLDSConnectionInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Connection Information";

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

        #endregion AD Forest Information

        #endregion AD Connection Information
    }
}
