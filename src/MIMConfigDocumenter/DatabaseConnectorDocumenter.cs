//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="DatabaseConnectorDocumenter.cs" company="Microsoft">
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
    /// The DatabaseConnectorDocumenter documents the configuration of an SQL / Oracle / DB2 connector.
    /// </summary>
    internal class DatabaseConnectorDocumenter : ConnectorDocumenter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DatabaseConnectorDocumenter"/> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="connectorName">The name.</param>
        /// <param name="configEnvironment">The environment in which the config element exists.</param>
        public DatabaseConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, ConfigEnvironment configEnvironment)
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
        /// Gets the database connector configuration report.
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

                this.ProcessDatabaseConnectionInformation();
                this.ProcessDatabaseColumnsConfiguration();
                this.ProcessConnectorFilterRules();
                this.ProcessConnectorJoinAndProjectionRules();
                this.ProcessConnectorAttributeFlows();
                this.ProcessConnectorDeprovisioningConfiguration();
                this.ProcessConnectorExtensionsConfiguration();
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

        #region Database Connection Information

        /// <summary>
        /// Processes the database connection information.
        /// </summary>
        protected void ProcessDatabaseConnectionInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Database Connection Information.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Setting, 3 = Configuration

                this.FillDatabaseConnectionInformationDataSet(true);
                this.FillDatabaseConnectionInformationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintDatabaseConnectionInformation();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the database connection information data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillDatabaseConnectionInformationDataSet(bool pilotConfig)
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

                    var serverName = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/connection-info/server");
                    var database = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/connection-info/databasename");
                    var dataSource = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/connection-info/datasource");
                    var tableName = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/connection-info/tablename");
                    var deltaTableName = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/connection-info/delta-tablename");
                    var multivalueTableName = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/connection-info/multivalued-tablename");
                    var integratedAuth = ((string)connector.XPathSelectElement("private-configuration/oledbma-configuration/connection-info/authentication") ?? string.Empty).Equals("integrated", StringComparison.OrdinalIgnoreCase);
                    var userName = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/connection-info/user");
                    var userDomain = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/connection-info/domain");

                    Documenter.AddRow(table, new object[] { 1, "Server Name", serverName });
                    if (!string.IsNullOrEmpty(dataSource))
                    {
                        Documenter.AddRow(table, new object[] { 2, "Data Source", dataSource });
                    }
                    else
                    {
                        Documenter.AddRow(table, new object[] { 2, "Database", database });
                    }

                    Documenter.AddRow(table, new object[] { 3, "Table/View", tableName });
                    Documenter.AddRow(table, new object[] { 4, "Delta Table/View", deltaTableName });
                    Documenter.AddRow(table, new object[] { 5, "Multivalue Table", multivalueTableName });

                    if (!this.ConnectorCategory.Equals("DB2", StringComparison.OrdinalIgnoreCase))
                    {
                        Documenter.AddRow(table, new object[] { 6, "Authentication", integratedAuth ? "Windows integrated authentication" : this.ConnectorCategory.Equals("MSSQL", StringComparison.OrdinalIgnoreCase) ? "SQL authentication" : "Database authentication" });
                    }

                    Documenter.AddRow(table, new object[] { 7, "User Name", userName });
                    Documenter.AddRow(table, new object[] { 8, "Password", "******" });
                    Documenter.AddRow(table, new object[] { 9, "Domain", userDomain });

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the database connection information.
        /// </summary>
        protected void PrintDatabaseConnectionInformation()
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

        #endregion Database Connection Information

        #region Database Columns Configuration

        /// <summary>
        /// Processes the database columns configuration.
        /// </summary>
        protected void ProcessDatabaseColumnsConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Database Columns Configuration.");

                var sectionTitle = "Columns Configuration";
                this.WriteSectionHeader(sectionTitle, 3);

                this.ProcessDatabaseColumnsInformation();
                this.ProcessDatabaseSpecialAttributes();
                this.ProcessDatabaseDeltaConfiguration();
                this.ProcessDatabaseMultiValueSettings();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Database Columns Information

        /// <summary>
        /// Processes the database connection information.
        /// </summary>
        protected void ProcessDatabaseColumnsInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Database Columns Information.");

                this.CreateSimpleSettingsDataSets(5); // 1 = Name, 2 = Database Type, 3 = Length, 4 = Nullable, 5 = Type

                this.FillDatabaseColumnsInformationDataSet(true);
                this.FillDatabaseColumnsInformationDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintDatabaseColumnsInformation();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the database columns information data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillDatabaseColumnsInformationDataSet(bool pilotConfig)
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

                    var columns = connector.XPathSelectElements("private-configuration/oledbma-configuration/mms-info/column-info/column");
                    foreach (var column in columns)
                    {
                        var name = (string)column.Element("name");
                        var dataType = (string)column.Element("data-type");
                        var length = (string)column.Element("length");
                        var nullable = (string)column.Element("isnullable") == "1" ? "Yes" : "No";
                        var type = column.Element("mms-type") != null && (string)column.Element("mms-type").Attribute("dn") == "1" ? "Reference DN" : (string)column.Element("mms-type");
                        Documenter.AddRow(table, new object[] { name, dataType, length, nullable, type });
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
        /// Prints the database columns information.
        /// </summary>
        protected void PrintDatabaseColumnsInformation()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Name", 30 }, { "Database Type", 30 }, { "Length", 10 }, { "Nullable", 10 }, { "Type", 20 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Database Columns Information

        #region Database Special Attributes

        /// <summary>
        /// Processes the database special attributes.
        /// </summary>
        protected void ProcessDatabaseSpecialAttributes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Database Special Attributes.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Setting, 3 = Configuration

                this.FillDatabaseSpecialAttributesDataSet(true);
                this.FillDatabaseSpecialAttributesDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintDatabaseSpecialAttributes();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the database columns information data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillDatabaseSpecialAttributesDataSet(bool pilotConfig)
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

                    var anchorAttributes = connector.XPathSelectElements("private-configuration/oledbma-configuration/mms-info/anchor/attribute");
                    var anchor = string.Empty;
                    foreach (var anchorAttribute in anchorAttributes)
                    {
                        anchor += (string)anchorAttribute + "+";
                    }

                    Documenter.AddRow(table, new object[] { 1, "Anchor", anchor.Trim('+') });

                    var objectTypeColumn = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/mms-info/object-type-info/object-type-column") ?? string.Empty;

                    if (string.IsNullOrEmpty(objectTypeColumn))
                    {
                        var objectType = (string)connector.XPathSelectElement("private-configuration/oledbma-configuration/mms-info/object-type");

                        Documenter.AddRow(table, new object[] { 2, "Fixed Object Type", (string)objectType });
                    }
                    else
                    {
                        Documenter.AddRow(table, new object[] { 3, "Object Type Column", (string)objectTypeColumn });

                        var objectTypes = connector.XPathSelectElements("private-configuration/oledbma-configuration/mms-info/object-type-info/object-type");
                        for (var objectTypeIndex = 0; objectTypeIndex < objectTypes.Count(); ++objectTypeIndex)
                        {
                            var objectType = (string)objectTypes.ElementAt(objectTypeIndex);

                            Documenter.AddRow(table, new object[] { 4 + objectTypeIndex, "Object Type", objectType });
                        }
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
        /// Prints the database columns information.
        /// </summary>
        protected void PrintDatabaseSpecialAttributes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Special Attributes", 30 }, { string.Empty, 70 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Database Special Attributes

        #region Database Delta Configuration

        /// <summary>
        /// Processes the database delta configuration.
        /// </summary>
        protected void ProcessDatabaseDeltaConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Database Delta Configuration.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Setting, 3 = Configuration

                this.FillDatabaseDeltaConfigurationDataSet(true);
                this.FillDatabaseDeltaConfigurationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintDatabaseDeltaConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the database delta configuration data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillDatabaseDeltaConfigurationDataSet(bool pilotConfig)
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

                    var deltaInfo = connector.XPathSelectElement("private-configuration/oledbma-configuration/mms-info/delta-info");

                    if (deltaInfo != null)
                    {
                        var changeColumn = (string)deltaInfo.Element("change-column");
                        var changeAdd = (string)deltaInfo.Element("add");
                        var changeModify = (string)deltaInfo.Element("update");
                        var changeDelete = (string)deltaInfo.Element("delete");
                        var attributeColumn = (string)deltaInfo.Element("attribute-column");
                        var attributeModify = (string)deltaInfo.Element("update-attribute");

                        Documenter.AddRow(table, new object[] { 1, "Change type attribute", changeColumn });
                        Documenter.AddRow(table, new object[] { 2, "Modify", changeModify });
                        Documenter.AddRow(table, new object[] { 3, "Add", changeAdd });
                        Documenter.AddRow(table, new object[] { 4, "Delete", changeDelete });
                        Documenter.AddRow(table, new object[] { 5, "Enable attribute-level change type synchronization", !string.IsNullOrEmpty(attributeColumn) ? "Yes" : "No" });

                        if (!string.IsNullOrEmpty(attributeColumn))
                        {
                            Documenter.AddRow(table, new object[] { 6, "Attribute-level synchronization column", attributeColumn });
                            Documenter.AddRow(table, new object[] { 7, "Attribute modify", attributeModify });
                        }
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
        /// Prints the database delta configuration.
        /// </summary>
        protected void PrintDatabaseDeltaConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Delta Configuration", 70 }, { string.Empty, 30 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Database Delta Configuration

        #region Database Multi-value Settings

        /// <summary>
        /// Processes the database Multi-value settings.
        /// </summary>
        protected void ProcessDatabaseMultiValueSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Database Multi-value Settings.");

                this.ProcessDatabaseMultiValueConfiguration();
                this.ProcessDatabaseMultiValueAttributeSettings();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Database Multi-value Configuration

        /// <summary>
        /// Processes the database Multi-value configuration.
        /// </summary>
        protected void ProcessDatabaseMultiValueConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Database Multi-value Configuration.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Setting, 3 = Configuration

                this.FillDatabaseMultiValueConfigurationDataSet(true);
                this.FillDatabaseMultiValueConfigurationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintDatabaseMultiValueConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the database multi-value data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillDatabaseMultiValueConfigurationDataSet(bool pilotConfig)
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

                    var multiValueInfo = connector.XPathSelectElement("private-configuration/oledbma-configuration/mms-info/multivalued-info");

                    if (multiValueInfo != null)
                    {
                        var attributeColumn = (string)multiValueInfo.Element("attribute-column");
                        var attributeColumnString = (string)multiValueInfo.Element("string-column");
                        var attributeColumnLargeString = (string)multiValueInfo.Element("long-string-column");
                        var attributeColumnBinary = (string)multiValueInfo.Element("binary-column");
                        var attributeColumnLargeBinary = (string)multiValueInfo.Element("long-binary-column");
                        var attributeColumnNumeric = (string)multiValueInfo.Element("numeric-column");

                        Documenter.AddRow(table, new object[] { 1, "Attribute name column", attributeColumn });
                        Documenter.AddRow(table, new object[] { 2, "String attribute column", attributeColumnString });
                        Documenter.AddRow(table, new object[] { 3, "Large string attribute column", attributeColumnLargeString });
                        Documenter.AddRow(table, new object[] { 4, "Binary attribute column", attributeColumnBinary });
                        Documenter.AddRow(table, new object[] { 5, "Large Binary attribute column", attributeColumnLargeBinary });
                        Documenter.AddRow(table, new object[] { 6, "Number attribute column", attributeColumnNumeric });
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
        /// Prints the database multi-value configuration.
        /// </summary>
        protected void PrintDatabaseMultiValueConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Multi-value Configuration", 70 }, { string.Empty, 30 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Database Delta Configuration

        #region Database Multi-value Attribute Settings

        /// <summary>
        /// Processes the database multi-value attribute settings.
        /// </summary>
        protected void ProcessDatabaseMultiValueAttributeSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Database Multi-value Attribute Settings.");

                this.CreateSimpleSettingsDataSets(2); // 1 = Name, 2 = Type

                this.FillDatabaseMultiValueAttributeSettingsDataSet(true);
                this.FillDatabaseMultiValueAttributeSettingsDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintDatabaseMultiValueAttributeSettings();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the database multi-value attribute settings data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillDatabaseMultiValueAttributeSettingsDataSet(bool pilotConfig)
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

                    var multiValueInfo = connector.XPathSelectElement("private-configuration/oledbma-configuration/mms-info/multivalued-info");

                    if (multiValueInfo != null)
                    {
                        var columns = connector.XPathSelectElements("private-configuration/oledbma-configuration/mms-info/multivalued-info/multivalued-columns/column");
                        foreach (var column in columns)
                        {
                            var name = (string)column.Element("name");
                            var type = column.Element("mms-type") != null && (string)column.Element("mms-type").Attribute("dn") == "1" ? "Reference DN" : (string)column.Element("mms-type");
                            Documenter.AddRow(table, new object[] { name, type });
                        }
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
        /// Prints the database multi-value attribute settings.
        /// </summary>
        protected void PrintDatabaseMultiValueAttributeSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Multi-value Attribute Settings", 50 }, { string.Empty, 50 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Database Multi-value Attribute Settings

        #endregion Database Multi-value Settings

        #endregion Database Columns Configuration
    }
}
