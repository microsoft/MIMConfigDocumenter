//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ConnectorDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// MIM Connector Configuration Documenter
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace MIMConfigDocumenter
{
    using System;
    using System.Collections.Specialized;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Web.UI;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The ConnectorDocumenter is an abstract base class for all types of MIM Sync Connectors.
    /// </summary>
    internal abstract class ConnectorDocumenter : Documenter
    {
        /// <summary>
        /// The logger context item connector name
        /// </summary>
        protected const string LoggerContextItemConnectorName = "Connector Name";

        /// <summary>
        /// The logger context item connector unique identifier
        /// </summary>
        protected const string LoggerContextItemConnectorGuid = "Connector Guid";

        /// <summary>
        /// The logger context item connector category
        /// </summary>
        protected const string LoggerContextItemConnectorCategory = "Connector Category";

        /// <summary>
        /// The logger context item connector sub type
        /// </summary>
        protected const string LoggerContextItemConnectorSubType = "Connector SubType";

        /// <summary>
        /// The data source object type currently being processed
        /// </summary>
        private string currentDataSourceObjectType;

        /// <summary>
        /// The metaverse object type currently being processed
        /// </summary>
        private string currentMetaverseObjectType;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConnectorDocumenter" /> class.
        /// </summary>
        /// <param name="pilotXml">The pilot configuration XML.</param>
        /// <param name="productionXml">The production configuration XML.</param>
        /// <param name="connectorName">The connector name.</param>
        /// <param name="configEnvironment">The environment in which the config element exists.</param>
        protected ConnectorDocumenter(XElement pilotXml, XElement productionXml, string connectorName, ConfigEnvironment configEnvironment)
        {
            Logger.Instance.WriteMethodEntry("Connector Name: '{0}'. Config Environment: '{1}'.", connectorName, configEnvironment);

            try
            {
                this.PilotXml = pilotXml;
                this.ProductionXml = productionXml;
                this.ConnectorName = connectorName;
                this.Environment = configEnvironment;

                string xpath = "//ma-data[name ='" + this.ConnectorName + "']";
                var connector = configEnvironment == ConfigEnvironment.ProductionOnly ? this.ProductionXml.XPathSelectElement(xpath, Documenter.NamespaceManager) : this.PilotXml.XPathSelectElement(xpath, Documenter.NamespaceManager);

                this.ConnectorGuid = (string)connector.Element("id");
                this.ConnectorCategory = (string)connector.Element("category");
                this.ConnectorSubType = (string)connector.Element("subtype");

                // Set Logger call context items
                Logger.SetContextItem(ConnectorDocumenter.LoggerContextItemConnectorName, this.ConnectorName);
                Logger.SetContextItem(ConnectorDocumenter.LoggerContextItemConnectorGuid, this.ConnectorGuid);
                Logger.SetContextItem(ConnectorDocumenter.LoggerContextItemConnectorCategory, this.ConnectorCategory);
                Logger.SetContextItem(ConnectorDocumenter.LoggerContextItemConnectorSubType, this.ConnectorSubType);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Connector Name: '{0}'. Config Environment: '{1}'.", connectorName, configEnvironment);
            }
        }

        /// <summary>
        /// Gets the name of the connector.
        /// </summary>
        /// <value>
        /// The name of the connector.
        /// </value>
        public string ConnectorName { get; private set; }

        /// <summary>
        /// Gets the connector unique identifier.
        /// </summary>
        /// <value>
        /// The connector unique identifier.
        /// </value>
        public string ConnectorGuid { get; private set; }

        /// <summary>
        /// Gets the connector category.
        /// </summary>
        /// <value>
        /// The connector category.
        /// </value>
        public string ConnectorCategory { get; private set; }

        /// <summary>
        /// Gets the type of the connector sub.
        /// </summary>
        /// <value>
        /// The type of the connector sub.
        /// </value>
        public string ConnectorSubType { get; private set; }

        /// <summary>
        /// Gets the type of the run profile step.
        /// </summary>
        /// <param name="runProfileStepType">Type of the run profile step.</param>
        /// <returns>The type of the run profile step.</returns>
        protected static string GetRunProfileStepType(XElement runProfileStepType)
        {
            Logger.Instance.WriteMethodEntry("Sync Rule Report Type: '{0}'.", runProfileStepType != null ? (string)runProfileStepType.Attribute("type") : null);

            var stepType = string.Empty;

            try
            {
                if (runProfileStepType != null)
                {
                    var type = ((string)runProfileStepType.Attribute("type") ?? string.Empty).ToUpperInvariant();
                    switch (type)
                    {
                        case "DELTA-IMPORT":
                            {
                                var importType = ((string)runProfileStepType.Element("import-subtype") ?? string.Empty).ToUpperInvariant();
                                stepType = importType == "TO-CS" ? "Delta Import (Stage Only)" : "Delta Import and Delta Synchronization";
                                break;
                            }

                        case "FULL-IMPORT":
                            {
                                var importType = ((string)runProfileStepType.Element("import-subtype") ?? string.Empty).ToUpperInvariant();
                                stepType = importType == "TO-CS" ? "Full Import (Stage Only)" : "Full Import and Delta Synchronization";
                                break;
                            }

                        case "EXPORT":
                            {
                                stepType = "Export";
                                break;
                            }

                        case "FULL-IMPORT-REEVALUATE-RULES":
                            {
                                stepType = "Full Import and Full Synchronization";
                                break;
                            }

                        case "APPLY-RULES":
                            {
                                var subType = ((string)runProfileStepType.Element("apply-rules-subtype") ?? string.Empty).ToUpperInvariant();
                                stepType = subType == "APPLY-PENDING" ? "Delta Synchronization" : subType == "REEVALUATE-FLOW-CONNECTORS" ? "Full Synchronization" : subType;
                                break;
                            }

                        default:
                            {
                                stepType = type;
                                break;
                            }
                    }
                }

                return stepType;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Sync Rule Report Type: '{0}'.", stepType);
            }
        }

        /// <summary>
        /// Writes the section header
        /// </summary>
        /// <param name="title">The section title</param>
        /// <param name="level">The section header level</param>
        protected override void WriteSectionHeader(string title, int level)
        {
            this.WriteSectionHeader(title, level, title, this.ConnectorGuid);
        }

        /// <summary>
        /// Writes the section header
        /// </summary>
        /// <param name="title">The section title</param>
        /// <param name="level">The section header level</param>
        /// <param name="bookmark">The section bookmark</param>
        protected new void WriteSectionHeader(string title, int level, string bookmark)
        {
            this.WriteSectionHeader(title, level, bookmark, this.ConnectorGuid);
        }

        /// <summary>
        /// Writes the connector report header.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        protected void WriteConnectorReportHeader()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

                var sectionTitle = this.ConnectorName + " Management Agent Configuration";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 2, this.ConnectorName, this.ConnectorGuid);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Connector Properties

        /// <summary>
        /// Processes the connector properties.
        /// </summary>
        protected void ProcessConnectorProperties()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Properties.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Property Name, 3 = Value

                this.FillConnectorPropertiesDataSet(true);
                this.FillConnectorPropertiesDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintConnectorProperties();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector properties data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorPropertiesDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var setting = (string)connector.Element("name");
                    Documenter.AddRow(table, new object[] { 0, "Connector Name", setting });

                    setting = (string)connector.Element("category");
                    Documenter.AddRow(table, new object[] { 1, "Connector Type", setting });

                    setting = (string)connector.Element("description");
                    Documenter.AddRow(table, new object[] { 2, "Description", setting });

                    setting = (string)connector.Element("subtype");
                    if (!string.IsNullOrEmpty(setting))
                    {
                        Documenter.AddRow(table, new object[] { 3, "Sub Type", setting });
                    }

                    setting = (string)connector.Element("ma-listname");
                    if (!string.IsNullOrEmpty(setting))
                    {
                        Documenter.AddRow(table, new object[] { 4, "List Name", setting });
                    }

                    setting = (string)connector.Element("ma-companyname");
                    if (!string.IsNullOrEmpty(setting))
                    {
                        Documenter.AddRow(table, new object[] { 5, "Company", setting });
                    }

                    setting = (string)connector.Element("creation-time");
                    Documenter.AddRow(table, new object[] { 6, "Creation Time", setting });

                    setting = (string)connector.Element("last-modification-time");
                    Documenter.AddRow(table, new object[] { 7, "Last Modification Time", setting });

                    setting = (string)connector.XPathSelectElement("controller-configuration/application-architecture");
                    Documenter.AddRow(table, new object[] { 8, "Architecture", setting });

                    setting = (string)connector.XPathSelectElement("controller-configuration/application-protection") == "low" ? "No" : "Yes";
                    Documenter.AddRow(table, new object[] { 9, "Run in Separate Process", setting });

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the connector properties.
        /// </summary>
        protected void PrintConnectorProperties()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Properties";

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

        #endregion Connector Properties

        #region Provisioning Hierarchy

        /// <summary>
        /// Processes the connector provisioning hierarchy configuration.
        /// </summary>
        protected void ProcessConnectorProvisioningHierarchyConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Provisioning Hierarchy Configuration.");

                this.CreateSimpleSettingsDataSets(2); // 1 = DN Component, 2 = Object Class Mapping

                this.FillConnectorProvisioningHierarchyConfigurationDataSet(true);
                this.FillConnectorProvisioningHierarchyConfigurationDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintConnectorProvisioningHierarchyConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector provisioning hierarchy configuration data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorProvisioningHierarchyConfigurationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var mappings = config.XPathSelectElements("//ma-data[name ='" + this.ConnectorName + "']/component_mappings/mapping");

                foreach (var mapping in mappings)
                {
                    Documenter.AddRow(table, new object[] { (string)mapping.Element("dn_component"), (string)mapping.Element("object_class") });
                }

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the connector provisioning hierarchy configuration.
        /// </summary>
        protected void PrintConnectorProvisioningHierarchyConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Provisioning Hierarchy";

                this.WriteSectionHeader(sectionTitle, 3);

                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "DN Component", 50 }, { "Object Class Mapping", 50 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
                else
                {
                    this.WriteContentParagraph("The provisioning hierarchy is not enabled.");
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Provisioning Hierarchy

        #region Container Selections

        /// <summary>
        /// Processes the connector partitions.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        protected void ProcessConnectorPartitionContainers(string partitionName)
        {
            this.CreateConnectorPartitionContainersDataSets();

            this.FillConnectorPartitionContainersDataSet(partitionName, true);
            this.FillConnectorPartitionContainersDataSet(partitionName, false);

            this.CreateConnectorPartitionContainersDiffgram();

            this.PrintConnectorPartitionContainers();
        }

        /// <summary>
        /// Creates the connector partition containers data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateConnectorPartitionContainersDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("Containers") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Container"); // Container
                var column2 = new DataColumn("Setting"); // Include / Exclude
                var column3 = new DataColumn("DNPart1"); // Sort Column 1
                var column4 = new DataColumn("DNPart2"); // Sort Column 2
                var column5 = new DataColumn("DNPart3"); // Sort Column 3
                var column6 = new DataColumn("DNPart4"); // Sort Column 4
                var column7 = new DataColumn("DNPart5"); // Sort Column 5
                var column8 = new DataColumn("DNPart6"); // Sort Column 6
                var column9 = new DataColumn("DNPart7"); // Sort Column 7
                var column10 = new DataColumn("DNPart8"); // Sort Column 8
                var column11 = new DataColumn("DNPart9"); // Sort Column 9
                var column12 = new DataColumn("DNPart10"); // Sort Column 10

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.Columns.Add(column4);
                table.Columns.Add(column5);
                table.Columns.Add(column6);
                table.Columns.Add(column7);
                table.Columns.Add(column8);
                table.Columns.Add(column9);
                table.Columns.Add(column10);
                table.Columns.Add(column11);
                table.Columns.Add(column12);
                table.PrimaryKey = new[] { column1, column2 };

                this.PilotDataSet = new DataSet("Containers") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetConnectorPartitionContainersPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector partition containers print table.
        /// </summary>
        /// <returns>The connector partition containers print table</returns>
        protected DataTable GetConnectorPartitionContainersPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Container
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Include / Exclude
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Sort Column1
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column2
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 3 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column3
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 4 }, { "Hidden", true }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column4
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 5 }, { "Hidden", true }, { "SortOrder", 3 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column5
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 6 }, { "Hidden", true }, { "SortOrder", 4 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column6
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 7 }, { "Hidden", true }, { "SortOrder", 5 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column7
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 8 }, { "Hidden", true }, { "SortOrder", 6 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column8
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 9 }, { "Hidden", true }, { "SortOrder", 7 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column9
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 10 }, { "Hidden", true }, { "SortOrder", 8 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                // Sort Column10
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 11 }, { "Hidden", true }, { "SortOrder", 9 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector partition containers data set.
        /// </summary>
        /// <param name="partitionName">Name of the partition.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorPartitionContainersDataSet(string partitionName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Partion Name: '{0}'. Pilot Config: '{1}'.", partitionName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var partition = connector.XPathSelectElement("ma-partition-data/partition[selected = 1 and name = '" + partitionName + "']");

                    if (partition != null)
                    {
                        var table = dataSet.Tables[0];

                        var inclusions = partition.XPathSelectElements("filter/containers/inclusions/inclusion");

                        var columnCount = table.Columns.Count;
                        foreach (var inclusion in inclusions)
                        {
                            var distinguishedName = (string)inclusion;
                            if (!string.IsNullOrEmpty(distinguishedName))
                            {
                                var row = this.GetContainerSelectionRow(distinguishedName, true, columnCount);
                                Documenter.AddRow(table, row);
                            }
                        }

                        var exclusions = partition.XPathSelectElements("filter/containers/exclusions/exclusion");

                        foreach (var exclusion in exclusions)
                        {
                            var distinguishedName = (string)exclusion;
                            if (!string.IsNullOrEmpty(distinguishedName))
                            {
                                var row = this.GetContainerSelectionRow(distinguishedName, false, columnCount);
                                Documenter.AddRow(table, row);
                            }
                        }

                        table.AcceptChanges();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Partion Name: '{0}'. Pilot Config: '{1}'.", partitionName, pilotConfig);
            }
        }

        /// <summary>
        /// Gets the container selection row.
        /// </summary>
        /// <param name="distinguishedName">The distinguished name of the container.</param>
        /// <param name="inclusion">True if the container is included.</param>
        /// <param name="columnCount">The column count.</param>
        /// <returns>The container selection row</returns>
        protected object[] GetContainerSelectionRow(string distinguishedName, bool inclusion, int columnCount)
        {
            Logger.Instance.WriteMethodEntry("Container: '{0}'. Included: '{1}'.", distinguishedName, inclusion);

            try
            {
                var row = new object[columnCount];
                var distinguishedNameParts = distinguishedName.Split(new string[] { "OU=" }, StringSplitOptions.None);
                var partsCount = distinguishedNameParts.Length;
                row[0] = distinguishedName;
                row[1] = inclusion ? "Include" : "Exclude";

                if (partsCount > Documenter.MaxSortableColumns)
                {
                    Logger.Instance.WriteInfo("Container: '{0}' is deeper than '{1}' levels. Display sequence may be a little out-of-order.", distinguishedName, Documenter.MaxSortableColumns);
                }

                for (var i = 0; i < row.Length - 2 && i < Documenter.MaxSortableColumns; ++i)
                {
                    row[2 + i] = string.Empty;
                    if (i < partsCount)
                    {
                        if (partsCount == 1)
                        {
                            row[2 + i] = " " + distinguishedNameParts[0]; // so that the domain root is always sorted first
                        }
                        else
                        {
                            row[2 + i] = distinguishedNameParts[partsCount - 1 - i];
                        }
                    }
                }

                return row;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Container: '{0}'. Included: '{1}'.", distinguishedName, inclusion);
            }
        }

        /// <summary>
        /// Creates the connector partition containers diffgram.
        /// </summary>
        protected void CreateConnectorPartitionContainersDiffgram()
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
        /// Prints the connector partition containers.
        /// </summary>
        protected void PrintConnectorPartitionContainers()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Container", 70 }, { "Include / Exclude", 30 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Container Selections

        #region Selected Object Types

        /// <summary>
        /// Processes the selected object types.
        /// </summary>
        protected void ProcessConnectorSelectedObjectTypes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Selected Object Types.");

                this.CreateSimpleSettingsDataSets(1);

                this.FillConnectorSelectedObjectTypesDataSet(true);
                this.FillConnectorSelectedObjectTypesDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintConnectorSelectedObjectTypes();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the selected object types data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorSelectedObjectTypesDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var objectTypes = config.XPathSelectElements("//ma-data[name ='" + this.ConnectorName + "']/ma-partition-data/partition[position() = 1]/filter/object-classes/object-class");

                foreach (var objectType in objectTypes)
                {
                    Documenter.AddRow(table, new object[] { (string)objectType });
                }

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the selected object types.
        /// </summary>
        protected void PrintConnectorSelectedObjectTypes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Selected Object Types";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Object Types", 100 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Selected Object Types

        #region Selected Attributes

        /// <summary>
        /// Processes the connector selected attributes.
        /// </summary>
        protected void ProcessConnectorSelectedAttributes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Selected Attributes. This may take a few minutes...");

                this.CreateSimpleSettingsDataSets(4);

                this.FillConnectorSelectedAttributesDataSet(true);
                this.FillConnectorSelectedAttributesDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintConnectorSelectedAttributes();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector selected attributes data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorSelectedAttributesDataSet(bool pilotConfig)
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
                    var currentConnectorGuid = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant(); // This may be pilot or production GUID

                    foreach (var attribute in connector.XPathSelectElements("attribute-inclusion/attribute"))
                    {
                        var attributeName = (string)attribute;

                        var attributeInfo = connector.XPathSelectElement(".//dsml:attribute-type[dsml:name = '" + attributeName + "']", Documenter.NamespaceManager);
                        if (attributeInfo != null)
                        {
                            var hasImportFlows = config.XPathSelectElement("//mv-data/import-attribute-flow/import-flow-set/import-flows/import-flow[translate(@src-ma, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + currentConnectorGuid + "' and */src-attribute = '" + attributeName + "']") != null;
                            var hasExportFlows = connector.XPathSelectElement("export-attribute-flow/export-flow-set/export-flow[@cd-attribute =  '" + attributeName + "']") != null;

                            var row = table.NewRow();

                            row[0] = attributeName;
                            var attributeSyntax = (string)attributeInfo.Element(Documenter.DsmlNamespace + "syntax");
                            row[1] = Documenter.GetAttributeType(attributeSyntax, (string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexable"));
                            row[2] = ((string)attributeInfo.Attribute("single-value") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes";

                            row[3] = hasImportFlows && hasExportFlows ? "Import / Export" : hasImportFlows ? "Import" : hasExportFlows ? "Export" : "No";

                            Documenter.AddRow(table, row);
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
        /// Prints the connector selected attributes.
        /// </summary>
        protected void PrintConnectorSelectedAttributes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Selected Attributes";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Attribute Name", 40 }, { "Type", 40 }, { "Multi-valued", 10 }, { "Flows Configured?", 10 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Selected Atrributes

        #region Connector Filter Rules

        /// <summary>
        /// Processes the connector filter rules.
        /// </summary>
        protected void ProcessConnectorFilterRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Filter Rules.");

                this.CreateConnectorFilterRulesDataSets();

                this.FillConnectorFilterRulesDataSet(true);
                this.FillConnectorFilterRulesDataSet(false);

                this.CreateConnectorFilterRulesDiffgram();

                this.PrintConnectorFilterRules();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the metaverse object type data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateConnectorFilterRulesDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("ConnectorFilter") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("SourceObjectType");
                var column2 = new DataColumn("FilterType");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("ConnectorFilterRuleGroup") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("SourceObjectType");
                var column22 = new DataColumn("RuleNumber", typeof(int));

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.PrimaryKey = new[] { column12, column22 };

                var table3 = new DataTable("ConnectorFilterRuleCondition") { Locale = CultureInfo.InvariantCulture };

                var column13 = new DataColumn("SourceObjectType");
                var column23 = new DataColumn("RuleNumber", typeof(int));
                var column33 = new DataColumn("SourceAttribute");
                var column43 = new DataColumn("Operator");
                var column53 = new DataColumn("Value");

                table3.Columns.Add(column13);
                table3.Columns.Add(column23);
                table3.Columns.Add(column33);
                table3.Columns.Add(column43);
                table3.Columns.Add(column53);
                table3.PrimaryKey = new[] { column13, column23, column33, column43, column53 };

                this.PilotDataSet = new DataSet("ConnectorFilterRules") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);
                this.PilotDataSet.Tables.Add(table3);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);
                var dataRelation23 = new DataRelation("DataRelation23", new[] { column12, column22 }, new[] { column13, column23 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);
                this.PilotDataSet.Relations.Add(dataRelation23);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetConnectorFilterRulesPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the metaverse object type print table.
        /// </summary>
        /// <returns>The metaverse object type print table.</returns>
        protected DataTable GetConnectorFilterRulesPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Source Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Filter Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Source Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Rule Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 3
                // Source Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Rule Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Source Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Operator
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", 3 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Value
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", 4 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector filter rules data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorFilterRulesDataSet(bool pilotConfig)
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
                    var table2 = dataSet.Tables[1];
                    var table3 = dataSet.Tables[2];

                    var filterRules = connector.XPathSelectElements("stay-disconnector/filter-set");

                    if (filterRules.Count() == 0)
                    {
                        return;
                    }

                    foreach (var filterRule in filterRules)
                    {
                        var objectType = (string)filterRule.Attribute("cd-object-type");
                        var filterType = (string)filterRule.Attribute("type");

                        if (string.IsNullOrEmpty(objectType))
                        {
                            continue;
                        }

                        Documenter.AddRow(table, new object[] { objectType, filterType });

                        var filterConditons = filterRule.XPathSelectElements("filter-alternative");
                        for (var filterConditionIndex = 0; filterConditionIndex < filterConditons.Count(); ++filterConditionIndex)
                        {
                            var filterConditon = filterConditons.ElementAt(filterConditionIndex);

                            Documenter.AddRow(table2, new object[] { objectType, filterConditionIndex + 1 });

                            foreach (var condition in filterConditon.Elements("condition"))
                            {
                                var filterAttribute = (string)condition.Attribute("cd-attribute");
                                var filterOperator = (string)condition.Attribute("operator");
                                var filterValue = (string)condition.Element("value");

                                Documenter.AddRow(table3, new object[] { objectType, filterConditionIndex + 1, filterAttribute, filterOperator, filterValue });
                            }
                        }
                    }

                    table.AcceptChanges();
                    table2.AcceptChanges();
                    table3.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Creates the connector filter rules difference gram.
        /// </summary>
        protected void CreateConnectorFilterRulesDiffgram()
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
        /// Gets the connector filter rules header table.
        /// </summary>
        /// <returns>The connector filter rules header table.</returns>
        protected DataTable GetConnectorFilterRulesHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                // Header Row 1
                // Data Source Object Type
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Data Source Object Type" }, { "RowSpan", 2 }, { "ColSpan", 1 } }).Values.Cast<object>().ToArray());

                // Filter Type
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 1 }, { "ColumnName", "Filter Type" }, { "RowSpan", 2 }, { "ColSpan", 1 } }).Values.Cast<object>().ToArray());

                // Filter
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 2 }, { "ColumnName", "Filter" }, { "RowSpan", 1 }, { "ColSpan", 4 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // #
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "#" }, { "RowSpan", 1 }, { "ColSpan", 1 } }).Values.Cast<object>().ToArray());

                // Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 } }).Values.Cast<object>().ToArray());

                // Operator
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Operator" }, { "RowSpan", 1 }, { "ColSpan", 1 } }).Values.Cast<object>().ToArray());

                // Value
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 3 }, { "ColumnName", "Value" }, { "RowSpan", 1 }, { "ColSpan", 1 } }).Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the connector filter rules.
        /// </summary>
        protected void PrintConnectorFilterRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Connector Filter Rules";

                this.WriteSectionHeader(sectionTitle, 3);

                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = this.GetConnectorFilterRulesHeaderTable();

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
                else
                {
                    this.WriteContentParagraph("There are no connector filter rules configured.");
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Connector Filter Rules

        #region Join and Projection Rules

        /// <summary>
        /// Processes the connector join and projection rules.
        /// </summary>
        protected void ProcessConnectorJoinAndProjectionRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ProcessConnectorJoinRules();
                this.ProcessConnectorProjectionRules();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Join Rules

        /// <summary>
        /// Processes the connector join rules.
        /// </summary>
        protected void ProcessConnectorJoinRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Join Rules.");

                this.CreateConnectorJoinRulesDataSets();

                this.FillConnectorJoinRulesDataSet(true);
                this.FillConnectorJoinRulesDataSet(false);

                this.CreateConnectorJoinRulesDiffgram();

                this.PrintConnectorJoinRules();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the metaverse object type data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateConnectorJoinRulesDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("ConnectorJoinRules") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("SourceObjectType");
                var column2 = new DataColumn("RuleNumber", typeof(int));
                var column3 = new DataColumn("MetaverseObjectType");
                var column4 = new DataColumn("JoinResolution");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.Columns.Add(column4);
                table.PrimaryKey = new[] { column1, column2, column3 };

                var table2 = new DataTable("ConnectorJoinRuleConditions") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("SourceObjectType");
                var column22 = new DataColumn("RuleNumber", typeof(int));
                var column32 = new DataColumn("MetaverseObjectType");
                var column42 = new DataColumn("SourceAttribute");
                var column52 = new DataColumn("MappingType");
                var column62 = new DataColumn("MetaverseAttribute");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.Columns.Add(column42);
                table2.Columns.Add(column52);
                table2.Columns.Add(column62);
                table2.PrimaryKey = new[] { column12, column22, column32, column42, column52, column62 };

                this.PilotDataSet = new DataSet("ConnectorJoinRules") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1, column2, column3 }, new[] { column12, column22, column32 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetConnectorJoinRulesPrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the metaverse object type print table.
        /// </summary>
        /// <returns>The metaverse object type print table.</returns>
        protected DataTable GetConnectorJoinRulesPrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Source Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Rule Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Metaverse Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Join Resolution
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Source Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Rule Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Metaverse Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Source Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", 3 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Mapping Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Metaverse Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 5 }, { "Hidden", false }, { "SortOrder", 4 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector join rules data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorJoinRulesDataSet(bool pilotConfig)
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
                    var table2 = dataSet.Tables[1];

                    var joinProfiles = from joinProfile in connector.XPathSelectElements("join/join-profile")
                                       let sourceObjectType = (string)joinProfile.Attribute("cd-object-type")
                                       orderby sourceObjectType
                                       select joinProfile;

                    if (joinProfiles.Count() == 0)
                    {
                        return;
                    }

                    var sourceObjectTypePrevious = "!Uninitialised!";
                    var mappingGroupIndex = 1;
                    for (var joinProfileIndex = 0; joinProfileIndex < joinProfiles.Count(); ++joinProfileIndex)
                    {
                        var joinProfile = joinProfiles.ElementAt(joinProfileIndex);
                        var sourceObjectType = (string)joinProfile.Attribute("cd-object-type");

                        var joinCriteria = joinProfile.Elements("join-criterion");

                        foreach (var joinCriterion in joinCriteria)
                        {
                            mappingGroupIndex = sourceObjectTypePrevious != sourceObjectType ? 1 : mappingGroupIndex + 1;
                            sourceObjectTypePrevious = sourceObjectType;
                            var metaverseObjectType = (string)joinCriterion.Element("search").Attribute("mv-object-type") ?? "Any";
                            var joinResoution = (string)joinCriterion.XPathSelectElement("resolution/script-context");

                            Documenter.AddRow(table, new object[] { sourceObjectType, mappingGroupIndex, metaverseObjectType, joinResoution });

                            var joinRuleType = (string)joinCriterion.Attribute("join-cri-type") ?? string.Empty;
                            foreach (var attributeMapping in joinCriterion.XPathSelectElements("search/attribute-mapping"))
                            {
                                var metaverseAttribute = (string)attributeMapping.Attribute("mv-attribute");
                                var sourceAttribute = (string)attributeMapping.XPathSelectElement("direct-mapping/src-attribute");

                                var mappingType = "Direct";
                                if (string.IsNullOrEmpty(sourceAttribute))
                                {
                                    foreach (var srcAttribute in attributeMapping.XPathSelectElements("scripted-mapping/src-attribute"))
                                    {
                                        sourceAttribute += (string)srcAttribute + ",";
                                    }

                                    sourceAttribute = sourceAttribute.TrimEnd(',');
                                    mappingType = "Rules Extension - " + (string)attributeMapping.XPathSelectElement("scripted-mapping/script-context");
                                }
                                else if (joinRuleType.Equals("sync-rule", StringComparison.OrdinalIgnoreCase))
                                {
                                    var scopeExpression = string.Empty;
                                    foreach (var scope in joinCriterion.XPathSelectElements("scoping/scope"))
                                    {
                                        scopeExpression += string.Format(CultureInfo.InvariantCulture, "{0} {1} {2} AND ", (string)scope.Element("csAttribute"), (string)scope.Element("csOperator"), (string)scope.Element("csValue"));
                                    }

                                    mappingType = string.IsNullOrEmpty(scopeExpression) ? "Sync Rule - Direct" : "Sync Rule - Scoped - " + scopeExpression.Substring(0, scopeExpression.Length - " AND ".Length);
                                }

                                Documenter.AddRow(table2, new object[] { sourceObjectType, mappingGroupIndex, metaverseObjectType, sourceAttribute, mappingType, metaverseAttribute });
                            }
                        }
                    }

                    table.AcceptChanges();
                    table2.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'", pilotConfig);
            }
        }

        /// <summary>
        /// Creates the connector join rules difference gram.
        /// </summary>
        protected void CreateConnectorJoinRulesDiffgram()
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
        /// Prints the connector join rules.
        /// </summary>
        protected void PrintConnectorJoinRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Join Rules";

                this.WriteSectionHeader(sectionTitle, 3);

                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Data Source Object Type", 15 }, { "Mapping Group#", 10 }, { "Metaverse Object Type", 15 }, { "Join Resolution", 15 }, { "Data Source Attribute", 15 }, { "Mapping Type", 15 }, { "Metaverse Attribute", 15 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
                else
                {
                    this.WriteContentParagraph("There are no join rules configured.");
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Join Rules

        #region Projetction Rules

        /// <summary>
        /// Processes the connector projection rules.
        /// </summary>
        protected void ProcessConnectorProjectionRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Projection Rules.");

                this.CreateSimpleOrderedSettingsDataSets(4, true); // 1 = Order Control, 2 = Data Source Object, 3 = Projection Type, 4 = Metaverse Object

                this.FillConnectorProjectionRulesDataSet(true);
                this.FillConnectorProjectionRulesDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintConnectorProjectionRules();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector projection rules data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorProjectionRulesDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var projectionRules = from projectionRule in connector.XPathSelectElements("projection/class-mapping")
                                          let sourceObjectType = (string)projectionRule.Attribute("cd-object-type")
                                          orderby sourceObjectType
                                          select projectionRule;

                    if (projectionRules.Count() == 0)
                    {
                        return;
                    }

                    var projectionRuleIndex = 0; // This will be the relative rule number if there are more than one sync rule based projection rules.
                    var previoussourceObjectType = string.Empty;
                    foreach (var projectionRule in projectionRules)
                    {
                        var sourceObjectType = (string)projectionRule.Attribute("cd-object-type");

                        projectionRuleIndex = sourceObjectType == previoussourceObjectType ? projectionRuleIndex + 1 : 0;
                        previoussourceObjectType = sourceObjectType;

                        var projectionType = (string)projectionRule.Attribute("type") ?? string.Empty;
                        var scopeExpression = string.Empty;
                        if (projectionType.Equals("sync-rule", StringComparison.OrdinalIgnoreCase))
                        {
                            foreach (var scope in projectionRule.XPathSelectElements("scoping/scope"))
                            {
                                scopeExpression += string.Format(CultureInfo.InvariantCulture, "{0} {1} {2} AND ", (string)scope.Element("csAttribute"), (string)scope.Element("csOperator"), (string)scope.Element("csValue"));
                            }

                            if (scopeExpression.Length > 5)
                            {
                                scopeExpression = scopeExpression.Substring(0, scopeExpression.Length - " AND ".Length);
                            }
                        }

                        projectionType = projectionType.Equals("scripted", StringComparison.OrdinalIgnoreCase) ?
                            "Rules Extension" : projectionType.Equals("declared", StringComparison.OrdinalIgnoreCase) ?
                            "Declared" : projectionType.Equals("sync-rule", StringComparison.OrdinalIgnoreCase) ? string.IsNullOrEmpty(scopeExpression) ? "Sync Rule - Direct" : "Sync Rule - Scoped - " + scopeExpression : projectionType;

                        var metaverseObjectType = (string)projectionRule.Element("mv-object-type") ?? string.Empty;

                        Documenter.AddRow(table, new object[] { sourceObjectType + projectionRuleIndex, sourceObjectType, projectionType, metaverseObjectType });
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
        /// Prints the connector projection rules.
        /// </summary>
        protected void PrintConnectorProjectionRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Projection Rules";

                this.WriteSectionHeader(sectionTitle, 3);

                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Data Source Object", 35 }, { "Projection Type", 30 }, { "Metaverse Object", 35 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
                else
                {
                    this.WriteContentParagraph("There are no projection rules configured.");
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Projetction Rules

        #endregion Join and Projection Rules

        #region Attribute Flows

        /// <summary>
        /// Processes the connector import and export attribute flows.
        /// </summary>
        protected void ProcessConnectorAttributeFlows()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Attribute Flows";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3);

                var pilotConnectorId = ((string)this.PilotXml.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']/id") ?? string.Empty).ToUpperInvariant();
                var productionConnectorId = ((string)this.ProductionXml.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']/id") ?? string.Empty).ToUpperInvariant();

                const string ImportFlowSetXPath = "//mv-data/import-attribute-flow/import-flow-set";
                var pilotConnectorIdXPath = "translate(@src-ma, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + pilotConnectorId + "'";
                var productionConnectorIdXPath = "translate(@src-ma, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + productionConnectorId + "'";

                var pilotSourceObjectTypesXPath = ImportFlowSetXPath + "/import-flows/import-flow[" + pilotConnectorIdXPath + "]";
                var productionSourceObjectTypesXPath = ImportFlowSetXPath + "/import-flows/import-flow[" + productionConnectorIdXPath + "]";
                var pilotMetaverseObjectTypesXPath = ImportFlowSetXPath + "[count(import-flows/import-flow[ " + pilotConnectorIdXPath + "]) != 0]";
                var productionMetaverseObjectTypesXPath = ImportFlowSetXPath + "[count(import-flows/import-flow[ " + productionConnectorIdXPath + "]) != 0]";

                var pilotSourceObjectTypes = from importFlow in this.PilotXml.XPathSelectElements(pilotSourceObjectTypesXPath, Documenter.NamespaceManager)
                                             let sourceObjectType = (string)importFlow.Attribute("cd-object-type")
                                             select sourceObjectType;

                var pilotSourceObjectTypes2 = from exportFlowSet in this.PilotXml.XPathSelectElements("//ma-data[name ='" + this.ConnectorName + "']/export-attribute-flow/export-flow-set")
                                              let sourceObjectType = (string)exportFlowSet.Attribute("cd-object-type")
                                              select sourceObjectType;

                var productionSourceObjectTypes = from importFlow in this.ProductionXml.XPathSelectElements(productionSourceObjectTypesXPath, Documenter.NamespaceManager)
                                                  let sourceObjectType = (string)importFlow.Attribute("cd-object-type")
                                                  select sourceObjectType;

                var productionSourceObjectTypes2 = from exportFlowSet in this.ProductionXml.XPathSelectElements("//ma-data[name ='" + this.ConnectorName + "']/export-attribute-flow/export-flow-set")
                                                   let sourceObjectType = (string)exportFlowSet.Attribute("cd-object-type")
                                                   select sourceObjectType;

                var sourceObjectTypes = pilotSourceObjectTypes.Union(pilotSourceObjectTypes2).Union(productionSourceObjectTypes).Union(productionSourceObjectTypes2).Distinct().OrderBy(pilotSourceObjectType => pilotSourceObjectType);

                var pilotMetaverseObjectTypes = from importFlowSet in this.PilotXml.XPathSelectElements(pilotMetaverseObjectTypesXPath, Documenter.NamespaceManager)
                                                let metaverseObjectType = (string)importFlowSet.Attribute("mv-object-type")
                                                select metaverseObjectType;

                var pilotMetaverseObjectTypes2 = from exportFlowSet in this.PilotXml.XPathSelectElements("//ma-data[name ='" + this.ConnectorName + "']/export-attribute-flow/export-flow-set")
                                                 let metaverseObjectType = (string)exportFlowSet.Attribute("mv-object-type")
                                                 select metaverseObjectType;

                var productionMetaverseObjectTypes = from importFlowSet in this.ProductionXml.XPathSelectElements(productionMetaverseObjectTypesXPath, Documenter.NamespaceManager)
                                                     let name = (string)importFlowSet.Attribute("mv-object-type")
                                                     select name;

                var productionMetaverseObjectTypes2 = from exportFlowSet in this.ProductionXml.XPathSelectElements("//ma-data[name ='" + this.ConnectorName + "']/export-attribute-flow/export-flow-set")
                                                      let metaverseObjectType = (string)exportFlowSet.Attribute("mv-object-type")
                                                      select metaverseObjectType;

                var metaverseObjectTypes = pilotMetaverseObjectTypes.Union(pilotMetaverseObjectTypes2).Union(productionMetaverseObjectTypes).Union(productionMetaverseObjectTypes2).Distinct().OrderBy(metaverseObjectType => metaverseObjectType);

                var connectorHasFlowsConfigured = false;

                foreach (var sourceObjectType in sourceObjectTypes)
                {
                    this.currentDataSourceObjectType = sourceObjectType;

                    foreach (var metaverseObjectType in metaverseObjectTypes)
                    {
                        this.currentMetaverseObjectType = metaverseObjectType;

                        var pilotHasImportFlowsXPath = ImportFlowSetXPath + "[@mv-object-type = '" + this.currentMetaverseObjectType + "' and count(import-flows/import-flow[@cd-object-type = '" + this.currentDataSourceObjectType + "' and " + pilotConnectorIdXPath + "])]";
                        var productionHasImportFlowsXPath = ImportFlowSetXPath + "[@mv-object-type = '" + this.currentMetaverseObjectType + "' and count(import-flows/import-flow[@cd-object-type = '" + this.currentDataSourceObjectType + "' and " + productionConnectorIdXPath + "])]";
                        var exportFlowsXPath = "//ma-data[name ='" + this.ConnectorName + "']/export-attribute-flow/export-flow-set[@mv-object-type = '" + this.currentMetaverseObjectType + "' and @cd-object-type = '" + this.currentDataSourceObjectType + "']/export-flow";

                        // Ignore the source object type and metaverse object type pair if there are no flows configured
                        var pilotHasImportFlows = this.PilotXml.XPathSelectElement(pilotHasImportFlowsXPath) != null;
                        var productionHasImportFlows = this.ProductionXml.XPathSelectElement(productionHasImportFlowsXPath) != null;
                        var pilotHasExportFlows = this.PilotXml.XPathSelectElement(exportFlowsXPath) != null;
                        var productionHasExportFlows = this.ProductionXml.XPathSelectElement(exportFlowsXPath) != null;

                        if (!pilotHasImportFlows && !productionHasImportFlows && !pilotHasExportFlows && !productionHasExportFlows)
                        {
                            continue;
                        }

                        connectorHasFlowsConfigured = true;

                        var arrows = (pilotHasImportFlows || productionHasImportFlows ? "&#8594" : string.Empty) + (pilotHasExportFlows || productionHasExportFlows ? "&#8592;" : string.Empty);
                        var sectionTitle2 = this.currentDataSourceObjectType + arrows + this.currentMetaverseObjectType;

                        this.WriteSectionHeader(sectionTitle2, 4);

                        this.ProcessConnectorObjectTypeImportAttributeFlows();
                        this.ProcessConnectorObjectTypeExportAttributeFlows();
                    }
                }

                if (!connectorHasFlowsConfigured)
                {
                    this.WriteContentParagraph("There are no attribute flows configured.");
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Import Flows

        /// <summary>
        /// Processes the connector import attribute flows.
        /// </summary>
        protected void ProcessConnectorObjectTypeImportAttributeFlows()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Import Attribute Flows. This may take a few minutes...");

                this.CreateSimpleOrderedSettingsDataSets(7, true); // 1 = Order Control, 2 = Data Source Attribute, 3 = To, 4 = Metaverse Attribute, 5 = Mapping Type, 6 = Scoping Filter, 7 = Precedence Order

                this.FillConnectorObjectTypeImportAttributeFlowsDataSet(true);
                this.FillConnectorObjectTypeImportAttributeFlowsDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintConnectorObjectTypeImportAttributeFlows();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector import attribute flows data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorObjectTypeImportAttributeFlowsDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var connectorId = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant();
                    var connectorIdXPath = "translate(@src-ma, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connectorId + "'";
                    var importFlowsXPath = "//mv-data/import-attribute-flow/import-flow-set" + "[@mv-object-type = '" + this.currentMetaverseObjectType + "']/import-flows/import-flow[@cd-object-type = '" + this.currentDataSourceObjectType + "' and " + connectorIdXPath + "]";
                    var metaverseAttributes = from importFlow in connector.XPathSelectElements(importFlowsXPath)
                                              let metaverseAttribute = (string)importFlow.Parent.Attribute("mv-attribute")
                                              orderby metaverseAttribute
                                              select metaverseAttribute;

                    if (metaverseAttributes.Count() == 0)
                    {
                        return;
                    }

                    var importFlowRuleIndex = 0; // This will be the relative rule number if there are more than one inbound sync rule import flows for the same metaverse attribute.
                    var previousMetaverseAttribute = string.Empty;
                    foreach (var metaverseAttribute in metaverseAttributes.Distinct())
                    {
                        var allImportFlowsXPath = "//mv-data/import-attribute-flow/import-flow-set[@mv-object-type = '" + this.currentMetaverseObjectType + "']/import-flows[@mv-attribute = '" + metaverseAttribute + "']";
                        var precedenceType = config.XPathSelectElement(allImportFlowsXPath) != null ? (string)config.XPathSelectElement(allImportFlowsXPath).Attribute("type") : string.Empty;
                        var allImportFlows = config.XPathSelectElements(allImportFlowsXPath + "/import-flow");

                        var importFlowRank = 0;
                        for (var importFlowIndex = 0; importFlowIndex < allImportFlows.Count(); ++importFlowIndex)
                        {
                            var importFlow = allImportFlows.ElementAt(importFlowIndex);
                            var importFlowConnectorId = ((string)importFlow.Attribute("src-ma") ?? string.Empty).ToUpperInvariant();

                            importFlowRank = precedenceType.Equals("ranked", StringComparison.OrdinalIgnoreCase) ? importFlowRank + 1 : 0;

                            if (importFlowConnectorId != connectorId)
                            {
                                continue;
                            }

                            importFlowRuleIndex = metaverseAttribute == previousMetaverseAttribute ? importFlowRuleIndex + 1 : 0;
                            previousMetaverseAttribute = metaverseAttribute;

                            var dataSourceAttribute = string.Empty;
                            var mappingType = string.Empty;
                            var scopeExpression = string.Empty;

                            if (importFlow.XPathSelectElement("direct-mapping/src-attribute") != null)
                            {
                                dataSourceAttribute = importFlow.XPathSelectElement("direct-mapping/src-attribute").Value;
                                mappingType = "Direct";
                            }
                            else if (importFlow.XPathSelectElement("scripted-mapping/src-attribute") != null)
                            {
                                foreach (var sourceAttribute in importFlow.XPathSelectElements("scripted-mapping/src-attribute"))
                                {
                                    dataSourceAttribute += (string)sourceAttribute + "<br/>";
                                }

                                dataSourceAttribute = !string.IsNullOrEmpty(dataSourceAttribute) ? dataSourceAttribute.Substring(0, dataSourceAttribute.Length - "<br/>".Length) : dataSourceAttribute;
                                mappingType = "Rules Extension - " + importFlow.XPathSelectElement("scripted-mapping/script-context").Value;
                            }
                            else if (importFlow.XPathSelectElement("constant-mapping/constant-value") != null)
                            {
                                dataSourceAttribute = string.Empty;
                                mappingType = "Constant - " + importFlow.XPathSelectElement("constant-mapping/constant-value").Value;
                            }
                            else if (importFlow.XPathSelectElement("dn-part-mapping/dn-part") != null)
                            {
                                dataSourceAttribute = string.Empty;
                                mappingType = "DN Component - ( " + importFlow.XPathSelectElement("dn-part-mapping/dn-part").Value + " )";
                            }
                            else if (importFlow.XPathSelectElement("sync-rule-mapping") != null)
                            {
                                mappingType = (string)importFlow.XPathSelectElement("sync-rule-mapping").Attribute("mapping-type") ?? string.Empty;
                                if (mappingType.Equals("direct", StringComparison.OrdinalIgnoreCase))
                                {
                                    dataSourceAttribute = importFlow.XPathSelectElement("sync-rule-mapping/src-attribute").Value;
                                    mappingType = "Sync Rule - Direct";
                                }
                                else if (mappingType.Equals("constant", StringComparison.OrdinalIgnoreCase))
                                {
                                    dataSourceAttribute = importFlow.XPathSelectElement("sync-rule-mapping/sync-rule-value").Value;
                                    mappingType = "Sync Rule - Constant";
                                }
                                else
                                {
                                    foreach (var sourceAttribute in importFlow.XPathSelectElements("sync-rule-mapping/src-attribute"))
                                    {
                                        dataSourceAttribute += (string)sourceAttribute + "<br/>";
                                    }

                                    dataSourceAttribute = !string.IsNullOrEmpty(dataSourceAttribute) ? dataSourceAttribute.Substring(0, dataSourceAttribute.Length - "<br/>".Length) : dataSourceAttribute;
                                    mappingType = "Sync Rule - Expression";  // TODO: Print the Sync Rule Expression
                                }

                                var syncRuleId = (string)importFlow.XPathSelectElement("sync-rule-mapping").Attribute("sync-rule-id") ?? string.Empty;

                                foreach (var scope in connector.XPathSelectElements("join/join-profile/join-criterion[@sync-rule-id='" + syncRuleId + "']/scoping/scope"))
                                {
                                    scopeExpression += string.Format(CultureInfo.InvariantCulture, "{0} {1} {2} AND ", (string)scope.Element("csAttribute"), (string)scope.Element("csOperator"), (string)scope.Element("csValue"));
                                }

                                if (scopeExpression.Length > 5)
                                {
                                    scopeExpression = scopeExpression.Substring(0, scopeExpression.Length - " AND ".Length);
                                }
                            }

                            Documenter.AddRow(table, new object[] { metaverseAttribute + importFlowRuleIndex, dataSourceAttribute, "&#8594;", metaverseAttribute, mappingType, scopeExpression, precedenceType.Equals("ranked", StringComparison.OrdinalIgnoreCase) ? importFlowRank.ToString(CultureInfo.InvariantCulture) : precedenceType });
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
        /// Prints the connector import attribute flows.
        /// </summary>
        protected void PrintConnectorObjectTypeImportAttributeFlows()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Data Source Attribute", 30 }, { "To", 5 }, { "Metaverse Attribute", 15 }, { "Mapping Type", 20 }, { "Scoping Filter", 23 }, { "Precedence Order", 7 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, HtmlTableSize.Huge);

                    this.WriteBreakTag();
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Import Flows

        #region Export Flows

        /// <summary>
        /// Processes the connector export attribute flows.
        /// </summary>
        protected void ProcessConnectorObjectTypeExportAttributeFlows()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Export Attribute Flows.");

                this.CreateSimpleOrderedSettingsDataSets(7, true); // 1 = Display Order Control, 2 = Data Source Attribute, 3 = From, 4 = Metaverse Attribute, 5 = Mapping Type, 6 = Allow Null?, 7 = Initial Flow?

                this.FillConnectorObjectTypeExportAttributeFlowsDataSet(true);
                this.FillConnectorObjectTypeExportAttributeFlowsDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintConnectorObjectTypeExportAttributeFlows();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector export attribute flows data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorObjectTypeExportAttributeFlowsDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var exportFlowsXPath = "export-attribute-flow/export-flow-set[@mv-object-type = '" + this.currentMetaverseObjectType + "' and @cd-object-type = '" + this.currentDataSourceObjectType + "']/export-flow";
                    var exportFlows = from exportFlow in connector.XPathSelectElements(exportFlowsXPath)
                                      let dataSourceAttribute = (string)exportFlow.Attribute("cd-attribute")
                                      orderby dataSourceAttribute
                                      select exportFlow;

                    if (exportFlows.Count() == 0)
                    {
                        return;
                    }

                    var exportFlowRuleIndex = 0; // This will be the relative rule number if there are more than one outbound sync rule export flows for the same data source attribute.
                    var previousDataSourceAttribute = string.Empty;
                    for (var exportFlowIndex = 0; exportFlowIndex < exportFlows.Count(); ++exportFlowIndex)
                    {
                        var exportFlow = exportFlows.ElementAt(exportFlowIndex);
                        var dataSourceAttribute = (string)exportFlow.Attribute("cd-attribute");

                        exportFlowRuleIndex = dataSourceAttribute == previousDataSourceAttribute ? exportFlowRuleIndex + 1 : 0;
                        previousDataSourceAttribute = dataSourceAttribute;

                        var metaverseAttribute = string.Empty;
                        var mappingType = string.Empty;
                        var allowNull = !((string)exportFlow.Attribute("suppress-deletions") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : string.Empty;
                        var initialFlowOnly = string.Empty;

                        if (exportFlow.XPathSelectElement("direct-mapping/src-attribute") != null)
                        {
                            metaverseAttribute = exportFlow.XPathSelectElement("direct-mapping/src-attribute").Value;
                            mappingType = "Direct";
                        }
                        else if (exportFlow.XPathSelectElement("scripted-mapping/src-attribute") != null)
                        {
                            foreach (var sourceAttribute in exportFlow.XPathSelectElements("scripted-mapping/src-attribute"))
                            {
                                metaverseAttribute += (string)sourceAttribute + "<br/>";
                            }

                            metaverseAttribute = !string.IsNullOrEmpty(metaverseAttribute) ? metaverseAttribute.Substring(0, metaverseAttribute.Length - "<br/>".Length) : metaverseAttribute;
                            mappingType = "Rules Extension - " + exportFlow.XPathSelectElement("scripted-mapping/script-context").Value;
                        }
                        else if (exportFlow.XPathSelectElement("constant-mapping/constant-value") != null)
                        {
                            metaverseAttribute = string.Empty;
                            mappingType = "Constant - " + exportFlow.XPathSelectElement("constant-mapping/constant-value").Value;
                        }
                        else if (exportFlow.XPathSelectElement("dn-part-mapping/dn-part") != null)
                        {
                            metaverseAttribute = string.Empty;
                            mappingType = "DN Component - ( " + exportFlow.XPathSelectElement("dn-part-mapping/dn-part").Value + " )";
                        }
                        else if (exportFlow.XPathSelectElement("sync-rule-mapping") != null)
                        {
                            mappingType = (string)exportFlow.XPathSelectElement("sync-rule-mapping").Attribute("mapping-type") ?? string.Empty;
                            if (mappingType.Equals("direct", StringComparison.OrdinalIgnoreCase))
                            {
                                metaverseAttribute = exportFlow.XPathSelectElement("sync-rule-mapping/src-attribute").Value;
                                mappingType = "Sync Rule - Direct";
                            }
                            else if (mappingType.Equals("constant", StringComparison.OrdinalIgnoreCase))
                            {
                                metaverseAttribute = exportFlow.XPathSelectElement("sync-rule-mapping/sync-rule-value").Value;
                                mappingType = "Sync Rule - Constant";
                            }
                            else
                            {
                                foreach (var sourceAttribute in exportFlow.XPathSelectElements("sync-rule-mapping/src-attribute"))
                                {
                                    metaverseAttribute += (string)sourceAttribute + "<br/>";
                                }

                                metaverseAttribute = !string.IsNullOrEmpty(metaverseAttribute) ? metaverseAttribute.Substring(0, metaverseAttribute.Length - "<br/>".Length) : metaverseAttribute;
                                mappingType = "Sync Rule - Expression";  // TODO: Print the Sync Rule Expression
                            }

                            initialFlowOnly = ((string)exportFlow.XPathSelectElement("sync-rule-mapping").Attribute("initial-flow-only") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : string.Empty;
                        }

                        Documenter.AddRow(table, new object[] { dataSourceAttribute + exportFlowRuleIndex, dataSourceAttribute, "&#8592;", metaverseAttribute, mappingType, allowNull, initialFlowOnly });
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
        /// Prints the connector export attribute flows.
        /// </summary>
        protected void PrintConnectorObjectTypeExportAttributeFlows()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Data Source Attribute", 20 }, { "From", 5 }, { "Metaverse Attribute", 35 }, { "Mapping Type", 26 }, { "Allow Null", 7 }, { "Initial Flow Only", 7 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, HtmlTableSize.Huge);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Import Flows

        #endregion Attribute Flows

        #region Run Profiles

        /// <summary>
        /// Processes the connector run profiles.
        /// </summary>
        protected void ProcessConnectorRunProfiles()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Run Profiles.");

                var sectionTitle = "Run Profiles";

                this.WriteSectionHeader(sectionTitle, 3);

                var xpath = "//ma-data[name ='" + this.ConnectorName + "']/ma-run-data/run-configuration";

                var pilot = this.PilotXml.XPathSelectElements(xpath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(xpath, Documenter.NamespaceManager);

                var connectorHasRunProfilesConfigured = false;

                // Sort by name
                var pilotRunProfiles = from runProfile in pilot
                                       let name = (string)runProfile.Element("name")
                                       orderby name
                                       select name;

                foreach (var runProfile in pilotRunProfiles)
                {
                    connectorHasRunProfilesConfigured = true;
                    this.ProcessConnectorRunProfile(runProfile);
                }

                // Sort by name
                var productionRunProfiles = from runProfile in production
                                            let name = (string)runProfile.Element("name")
                                            orderby name
                                            select name;

                productionRunProfiles = productionRunProfiles.Where(productionRunProfile => !pilotRunProfiles.Contains(productionRunProfile));

                foreach (var runProfile in productionRunProfiles)
                {
                    connectorHasRunProfilesConfigured = true;
                    this.ProcessConnectorRunProfile(runProfile);
                }

                if (!connectorHasRunProfilesConfigured)
                {
                    this.WriteContentParagraph("There are no run profiles configured.");
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the connector run profile.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        protected void ProcessConnectorRunProfile(string runProfileName)
        {
            Logger.Instance.WriteMethodEntry("Run Profile Name: '{0}'.", runProfileName);

            try
            {
                this.WriteSectionHeader("Run Profile: " + runProfileName, 4, runProfileName);

                this.CreateConnectorRunProfileDataSets();

                this.FillConnectorRunProfileDataSet(runProfileName, true);
                this.FillConnectorRunProfileDataSet(runProfileName, false);

                this.CreateConnectorRunProfileDiffgram();

                this.PrintConnectorRunProfile();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Run Profile Name: '{0}'.", runProfileName);
            }
        }

        /// <summary>
        /// Creates the connector run profile data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateConnectorRunProfileDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("ConnectorRunProfiles") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Step Number", typeof(int));
                var column2 = new DataColumn("Step Name");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("ConnectorRunProfileConfiguration") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("Step Number", typeof(int));
                var column22 = new DataColumn("Setting");
                var column32 = new DataColumn("Configuration");
                var column42 = new DataColumn("Setting Number", typeof(int));

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.Columns.Add(column42);
                table2.PrimaryKey = new[] { column12, column22 };

                this.PilotDataSet = new DataSet("ConnectorRunProfileConfigurations") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetConnectorRunProfilePrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the connector run profile print table.
        /// </summary>
        /// <returns>
        /// The connector run profile print table.
        /// </returns>
        protected DataTable GetConnectorRunProfilePrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Step Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Step Name
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Step Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Setting
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Configuration
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Setting Number
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector run profile data set.
        /// </summary>
        /// <param name="runProfileName">Name of the run profile.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected virtual void FillConnectorRunProfileDataSet(string runProfileName, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Run Profile Name: '{0}'. Pilot Config: '{1}'.", runProfileName, pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var table = dataSet.Tables[0];
                    var table2 = dataSet.Tables[1];

                    var runProfileSteps = connector.XPathSelectElements("ma-run-data/run-configuration[name = '" + runProfileName + "']/configuration/step");

                    for (var stepIndex = 1; stepIndex <= runProfileSteps.Count(); ++stepIndex)
                    {
                        var runProfileStep = runProfileSteps.ElementAt(stepIndex - 1);

                        var runProfileStepType = ConnectorDocumenter.GetRunProfileStepType(runProfileStep.Element("step-type"));

                        Documenter.AddRow(table, new object[] { stepIndex, runProfileStepType });

                        var logFileName = (string)runProfileStep.Element("dropfile-name");
                        if (!string.IsNullOrEmpty(logFileName))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Log file", logFileName, 1 });
                        }

                        var numberOfObjects = (string)runProfileStep.XPathSelectElement("threshold/object");
                        if (!string.IsNullOrEmpty(numberOfObjects))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Number of objects", numberOfObjects, 2 });
                        }

                        var numberOfDeletions = (string)runProfileStep.XPathSelectElement("threshold/delete");
                        if (!string.IsNullOrEmpty(numberOfDeletions))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Number of deletions", numberOfDeletions, 3 });
                        }

                        var partitionId = ((string)runProfileStep.Element("partition") ?? string.Empty).ToUpperInvariant();
                        var partitionName = (string)connector.XPathSelectElement("ma-partition-data/partition[translate(id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + partitionId + "']/name");
                        Documenter.AddRow(table2, new object[] { stepIndex, "Partition", partitionName, 4 });

                        var inputFileName = (string)connector.XPathSelectElement("custom-data/run-config/input-file");
                        if (!string.IsNullOrEmpty(inputFileName))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Input file name", inputFileName, 5 });
                        }

                        var outputFileName = (string)connector.XPathSelectElement("custom-data/run-config/output-file");
                        if (!string.IsNullOrEmpty(outputFileName))
                        {
                            Documenter.AddRow(table2, new object[] { stepIndex, "Output file name", outputFileName, 6 });
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

        /// <summary>
        /// Creates the connector run profile diffgram.
        /// </summary>
        protected void CreateConnectorRunProfileDiffgram()
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
        /// Prints the extensible2 run profile.
        /// </summary>
        protected void PrintConnectorRunProfile()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Step#", 5 }, { "Step Name", 35 }, { "Setting", 35 }, { "Configuration", 25 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Run Profiles

        #region Deprovisioning

        /// <summary>
        /// Processes the connector deprovisioning configuration.
        /// </summary>
        protected void ProcessConnectorDeprovisioningConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Deprovisioning Configuration.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Property Name, 3 = Value

                this.FillConnectorDeprovisioningConfigurationDataSet(true);
                this.FillConnectorDeprovisioningConfigurationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintConnectorDeprovisioningConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector deprovisioning configuration data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorDeprovisioningConfigurationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var deprovisioningAction = (string)connector.XPathSelectElement("provisioning-cleanup/action") ?? string.Empty;

                    if (deprovisioningAction.Equals("make-normal-disconnector", StringComparison.OrdinalIgnoreCase))
                    {
                        deprovisioningAction = "Make them disconnectors";
                    }
                    else if (deprovisioningAction.Equals("make-explicit-disconnector", StringComparison.OrdinalIgnoreCase))
                    {
                        deprovisioningAction = "Make them explicit disconnectors";
                    }
                    else if (deprovisioningAction.Equals("delete-object", StringComparison.OrdinalIgnoreCase))
                    {
                        deprovisioningAction = "Stage a delete on the object for the next export run";
                    }
                    else
                    {
                        var deprovisioningType = connector.XPathSelectElement("provisioning-cleanup") != null ? (string)connector.XPathSelectElement("provisioning-cleanup").Attribute("type") : string.Empty;

                        if (deprovisioningType.Equals("scripted", StringComparison.OrdinalIgnoreCase))
                        {
                            deprovisioningAction = "Determine with a rules extension";
                        }
                    }

                    Documenter.AddRow(table, new object[] { 0, "Deprovisioning Option", deprovisioningAction });

                    var connectorId = ((string)connector.Element("id") ?? string.Empty).ToUpperInvariant();
                    var connectorIdXPath = "translate(@ma-id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connectorId + "'";
                    var enableAttributeRecall = ((string)config.XPathSelectElement("//mv-data/import-attribute-flow/per-ma-options/ma-options[" + connectorIdXPath + "]/enable-recall") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes";

                    Documenter.AddRow(table, new object[] { 1, "Do not recall attributes contributed by an object from this MA when it is disconnected", enableAttributeRecall });

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Prints the connector deprovisioning configuration.
        /// </summary>
        protected void PrintConnectorDeprovisioningConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Deprovisioning";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Setting", 70 }, { "Configuration", 30 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Deprovisioning

        #region Extension Configuration

        /// <summary>
        /// Processes the connector extensions configuration.
        /// </summary>
        protected void ProcessConnectorExtensionsConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Extensions Configuration.");

                this.ProcessConnectorRulesExtensionConfiguration();
                this.ProcessConnectorPasswordManagementConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Rules Extension Configuration

        /// <summary>
        /// Processes the connector rules extension configuration.
        /// </summary>
        protected void ProcessConnectorRulesExtensionConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Rules Extension Configuration.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Property Name, 3 = Value

                this.FillConnectorRulesExtensionConfigurationDataSet(true);
                this.FillConnectorRulesExtensionConfigurationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintConnectorRulesExtensionConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector rules extension configuration data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorRulesExtensionConfigurationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var assemblyName = (string)connector.XPathSelectElement("extension/assembly-name");

                    if (!string.IsNullOrEmpty(assemblyName))
                    {
                        Documenter.AddRow(table, new object[] { 0, "Rules extension name", assemblyName });

                        var runInSeparateProcess = !((string)connector.XPathSelectElement("extension/application-protection") ?? string.Empty).Equals("low", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No";
                        Documenter.AddRow(table, new object[] { 1, "Run this rules extension in a separate process", runInSeparateProcess });
                    }
                    else
                    {
                        Documenter.AddRow(table, new object[] { 0, "Rules extension name", "-" });
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
        /// Prints the connector rules extension configuration.
        /// </summary>
        protected void PrintConnectorRulesExtensionConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Extensions Configuration";

                this.WriteSectionHeader(sectionTitle, 3);

                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Setting", 50 }, { "Configuration", 50 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);

                this.WriteBreakTag();
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Rules Extension Configuration

        #region Password Management Extension Configuration

        /// <summary>
        /// Processes the connector password management configuration.
        /// </summary>
        protected void ProcessConnectorPasswordManagementConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Connector Password Management Configuration.");

                this.CreateSimpleOrderedSettingsDataSets(3); // 1 = Display Order Control, 2 = Property Name, 3 = Value

                this.FillConnectorPasswordManagementConfigurationDataSet(true);
                this.FillConnectorPasswordManagementConfigurationDataSet(false);

                this.CreateSimpleOrderedSettingsDiffgram();

                this.PrintConnectorPasswordManagementConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector password management configuration data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorPasswordManagementConfigurationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    var enablePasswordSync = !((string)connector.XPathSelectElement("password-sync-allowed") ?? string.Empty).Equals("1", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes";

                    Documenter.AddRow(table, new object[] { 0, "Enable password management", enablePasswordSync });

                    if (enablePasswordSync.Equals("Yes", StringComparison.OrdinalIgnoreCase))
                    {
                        if (connector.XPathSelectElement("private-configuration/*/password-extension-config[password-extension-enabled = '1']") != null)
                        {
                            Documenter.AddRow(table, new object[] { 1, "Extension name", (string)connector.XPathSelectElement("private-configuration/*/password-extension-config/dll") });

                            var connectionInfo = connector.XPathSelectElement("private-configuration/*/password-extension-config/connection-info");

                            if (connectionInfo != null)
                            {
                                Documenter.AddRow(table, new object[] { 2, "Connect To", (string)connectionInfo.Element("connect-to") });
                                Documenter.AddRow(table, new object[] { 3, "User", (string)connectionInfo.Element("user") });
                                Documenter.AddRow(table, new object[] { 4, "Password", "******" });
                            }

                            Documenter.AddRow(table, new object[] { 5, "Connection timeout (seconds)", (string)connector.XPathSelectElement("private-configuration/*/password-extension-config/timeout") });

                            var passwordSetEnabled = ((string)connector.XPathSelectElement("private-configuration/*/password-extension-config/password-set-enabled")).Equals("1", StringComparison.OrdinalIgnoreCase);
                            var passwordChangeEnabled = ((string)connector.XPathSelectElement("private-configuration/*/password-extension-config/password-change-enabled")).Equals("1", StringComparison.OrdinalIgnoreCase);
                            Documenter.AddRow(table, new object[] { 6, "Supported password operations", passwordSetEnabled && passwordChangeEnabled ? "Set and change" : passwordSetEnabled ? "Set" : passwordChangeEnabled ? "Change" : string.Empty });
                        }

                        // Password sync target settings
                        Documenter.AddRow(table, new object[] { 11, "Maximum retry count", (string)connector.XPathSelectElement("password-sync/maximum-retry-count") });
                        Documenter.AddRow(table, new object[] { 12, "Retry interval (seconds)", (string)connector.XPathSelectElement("password-sync/retry-interval") });
                        Documenter.AddRow(table, new object[] { 13, "Require secure connection for password synchronization options", !((string)connector.XPathSelectElement("password-sync/allow-low-security") ?? string.Empty).Equals("1", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No" });
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
        /// Prints the connector password management configuration.
        /// </summary>
        protected void PrintConnectorPasswordManagementConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Password Management Setting", 70 }, { "Configuration", 30 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);

                this.WriteBreakTag();
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Password Management Extension Configuration

        #endregion Extension Configuration

        #region Partition Display Names Configuration

        /// <summary>
        /// Processes the connector partition display names configuration.
        /// </summary>
        protected void ProcessConnectorPartitionDisplayNamesConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Partition Display Names Configuration.");

                this.CreateSimpleSettingsDataSets(2); // 1= Partition Name, 3 = Display Name

                this.FillConnectorPartitionDisplayNamesConfigurationDataSet(true);
                this.FillConnectorPartitionDisplayNamesConfigurationDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintConnectorPartitionDisplayNamesConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the connector partition display names data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillConnectorPartitionDisplayNamesConfigurationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var connector = config.XPathSelectElement("//ma-data[name ='" + this.ConnectorName + "']");

                if (connector != null)
                {
                    foreach (var partition in connector.XPathSelectElements("ma-partition-data/partition[selected > 0]"))
                    {
                        var name = (string)partition.Element("name");
                        var displayName = (string)partition.Element("display-name");
                        Documenter.AddRow(table, new object[] { name, displayName });
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
        /// Prints the connector partition display names configuration.
        /// </summary>
        protected void PrintConnectorPartitionDisplayNamesConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Partition", 50 }, { "Display Name", 50 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);

                    this.WriteBreakTag();
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Partition Display Names Configuration
    }
}
