//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ActiveDirectoryGALConnectorDocumenter.cs" company="Microsoft">
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
    /// The ActiveDirectoryGALConnectorDocumenter documents the configuration of an Active Directory GAL connector.
    /// </summary>
    internal class ActiveDirectoryGALConnectorDocumenter : ActiveDirectoryConnectorDocumenter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ActiveDirectoryGALConnectorDocumenter"/> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="connectorName">The name.</param>
        /// <param name="configEnvironment">The environment in which the config element exists.</param>
        public ActiveDirectoryGALConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, ConfigEnvironment configEnvironment)
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

                this.ProcessActiveDirectoryConnectionInformation();
                this.ProcessActiveDirectoryPartitions();
                this.ProcessActiveDirectoryGALConfiguration();
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

        #region AD GAL Configuration

        /// <summary>
        /// Processes the AD GAL configuration.
        /// </summary>
        protected void ProcessActiveDirectoryGALConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Active Directory GAL Configuration.");

                var sectionTitle = "GAL Configuration";

                this.WriteSectionHeader(sectionTitle, 3);

                this.ProcessActiveDirectoryGALContainerConfiguration();
                this.ProcessActiveDirectoryGALExchangeConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region AD GAL Container Configuration

        /// <summary>
        /// Processes the AD password management configuration.
        /// </summary>
        protected void ProcessActiveDirectoryGALContainerConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Active Directory GAL Container Configuration.");

                this.CreateActiveDirectoryGALContainerConfigurationDataSets();

                this.FillActiveDirectoryGALContainerConfigurationDataSet(true);
                this.FillActiveDirectoryGALContainerConfigurationDataSet(false);

                this.CreateActiveDirectoryGALContainerConfigurationDiffgram();

                this.PrintActiveDirectoryGALContainerConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the connector GAL container configuration data set.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateActiveDirectoryGALContainerConfigurationDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("GALContainerConfigurationSetting") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("DisplayOrder", typeof(int));
                var column2 = new DataColumn("Setting");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("GALContainerConfiguration") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("DisplayOrder", typeof(int));
                var column22 = new DataColumn("Configuration");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.PrimaryKey = new[] { column12, column22 };

                this.PilotDataSet = new DataSet("GALContainerConfiguration") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetActiveDirectoryGALContainerConfigurationPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector GAL container configuration print table.
        /// </summary>
        /// <returns>The connector GAL container configuration print table.</returns>
        protected DataTable GetActiveDirectoryGALContainerConfigurationPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Display Order
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Setting
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Configuration
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector GAL container configuration data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillActiveDirectoryGALContainerConfigurationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];
                var table2 = dataSet.Tables[1];

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var targetOu = (string)connector.XPathSelectElement("private-configuration/adma-configuration/ui-data/galma/target-ou");

                    Documenter.AddRow(table, new object[] { 0, "Destination container for contacts synchronized with this forest" });
                    Documenter.AddRow(table2, new object[] { 0, targetOu });

                    var sourceOus = connector.XPathSelectElements("private-configuration/adma-configuration/ui-data/galma/source-contact-ous/ou");

                    Documenter.AddRow(table, new object[] { 1, "Source containers with authoritative contacts in the forest" });

                    if (sourceOus.Count() == 0)
                    {
                        Documenter.AddRow(table2, new object[] { 1, "None" });
                    }
                    else
                    {
                        foreach (var sourceOu in sourceOus)
                        {
                            Documenter.AddRow(table2, new object[] { 1, sourceOu.Value });
                        }
                    }

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Creates the connector GAL container configuration difference gram.
        /// </summary>
        protected void CreateActiveDirectoryGALContainerConfigurationDiffgram()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = Documenter.GetDiffgram(this.PilotDataSet, this.ProductionDataSet);
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector GAL container configuration.
        /// </summary>
        protected void PrintActiveDirectoryGALContainerConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "GAL Container Configuration", 50 }, { string.Empty, 50 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion AD GAL Container Configuration

        #region AD GAL Exchange Configuration

        /// <summary>
        /// Processes the AD GAL Exchange configuration.
        /// </summary>
        protected void ProcessActiveDirectoryGALExchangeConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Active Directory GAL Exchange Configuration.");

                this.CreateActiveDirectoryGALExchangeConfigurationDataSets();

                this.FillActiveDirectoryGALExchangeConfigurationDataSet(true);
                this.FillActiveDirectoryGALExchangeConfigurationDataSet(false);

                this.CreateActiveDirectoryGALExchangeConfigurationDiffgram();

                this.PrintActiveDirectoryGALExchangeConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the connector GAL Exchange configuration data set.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateActiveDirectoryGALExchangeConfigurationDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("GALExchangeConfigurationSetting") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("DisplayOrder", typeof(int));
                var column2 = new DataColumn("Setting");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("GALExchangeConfiguration") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("DisplayOrder", typeof(int));
                var column22 = new DataColumn("Configuration");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.PrimaryKey = new[] { column12, column22 };

                this.PilotDataSet = new DataSet("GALExchangeConfiguration") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetActiveDirectoryGALExchangeConfigurationPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector GAL Exchange configuration print table.
        /// </summary>
        /// <returns>The connector GAL Exchange configuration print table.</returns>
        protected DataTable GetActiveDirectoryGALExchangeConfigurationPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Display Order
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Setting
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Configuration
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector GAL Exchange configuration data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillActiveDirectoryGALExchangeConfigurationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];
                var table2 = dataSet.Tables[1];

                var connector = config.XPathSelectElement(Documenter.GetConnectorXmlRootXPath(pilotConfig) + "/ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var smtpDomains = connector.XPathSelectElements("private-configuration/adma-configuration/ui-data/galma/smtp-mail-domains/domain");

                    Documenter.AddRow(table, new object[] { 0, "SMTP mail suffix(s) for mailbox and mail objects in this forest" });
                    if (smtpDomains.Count() == 0)
                    {
                        Documenter.AddRow(table2, new object[] { 0, "None" });
                    }
                    else
                    {
                        foreach (var smtpDomain in smtpDomains)
                        {
                            Documenter.AddRow(table2, new object[] { 0, smtpDomain.Value });
                        }
                    }

                    var mailRouting = ((string)connector.XPathSelectElement("private-configuration/adma-configuration/ui-data/galma/mail-routing") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No";

                    Documenter.AddRow(table, new object[] { 1, "Route mail through this forest for all contacts created from contacts in this forest" });
                    Documenter.AddRow(table2, new object[] { 1, mailRouting });

                    var crossForestDelegation = ((string)connector.XPathSelectElement("private-configuration/adma-configuration/ui-data/galma/cross-forest-delegation") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No";

                    Documenter.AddRow(table, new object[] { 2, "Support cross-forest delegation (Exchange 2007 or 2010 only)" });
                    Documenter.AddRow(table2, new object[] { 2, crossForestDelegation });

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Creates the connector GAL Exchange configuration difference gram.
        /// </summary>
        protected void CreateActiveDirectoryGALExchangeConfigurationDiffgram()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = Documenter.GetDiffgram(this.PilotDataSet, this.ProductionDataSet);
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector GAL Exchange configuration.
        /// </summary>
        protected void PrintActiveDirectoryGALExchangeConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "GAL Exchange Configuration", 60 }, { string.Empty, 40 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion AD GAL Exchange Configuration

        #endregion AD GAL Configuration
    }
}
