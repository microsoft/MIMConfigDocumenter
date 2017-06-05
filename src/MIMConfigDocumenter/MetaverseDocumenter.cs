//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="MetaverseDocumenter.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// MIM Metaverse Configuration Documenter
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace MIMConfigDocumenter
{
    using System;
    using System.Collections.Generic;
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
    /// The MetaverseDocumenter documents the configuration of Metaverse.
    /// </summary>
    internal class MetaverseDocumenter : Documenter
    {
        /// <summary>
        /// The logger context item metaverse object type
        /// </summary>
        private const string LoggerContextItemMetaverseObjectType = "Metaverse Object Type";

        /// <summary>
        /// The logger context item metaverse attribute
        /// </summary>
        private const string LoggerContextItemMetaverseAttribute = "Metaverse Attribute";

        /// <summary>
        /// The object type currently being processed
        /// </summary>
        private string currentObjectType;

        /// <summary>
        /// Initializes a new instance of the <see cref="MetaverseDocumenter"/> class.
        /// </summary>
        /// <param name="pilotConfigXml">The pilot configuration XML.</param>
        /// <param name="productionConfigXml">The production configuration XML.</param>
        public MetaverseDocumenter(XElement pilotConfigXml, XElement productionConfigXml)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PilotXml = pilotConfigXml;
                this.ProductionXml = productionConfigXml;

                this.ReportFileName = Documenter.GetTempFilePath("MV.tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath("MV.TOC.tmp.html");
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the metaverse configuration report.
        /// </summary>
        /// <returns>The Tuple of configuration report and associated TOC</returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

                var sectionTitle = "Metaverse Configuration";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 2);

                this.ProcessMetaverseObjectTypes();
                this.ProcessMetaverseObjectDeletionRules();
                this.ProcessMetaverseOptions();

                return this.GetReportTuple();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Metaverse ObjectType

        /// <summary>
        /// Processes the metaverse object types.
        /// </summary>
        private void ProcessMetaverseObjectTypes()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Metaverse Object Types";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3);

                const string XPath = "//mv-data//dsml:class";
                var pilot = this.PilotXml.XPathSelectElements(XPath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(XPath, Documenter.NamespaceManager);

                // Sort by name
                var pilotObjectTypes = from objectType in pilot
                                       let name = (string)objectType.Element(Documenter.DsmlNamespace + "name")
                                       orderby name
                                       select name;

                foreach (var objectType in pilotObjectTypes)
                {
                    this.currentObjectType = objectType;

                    // Ignore the object type if there are no import flows configured
                    var pilotHasImportFlows = this.PilotXml.XPathSelectElement("//mv-data/import-attribute-flow/import-flow-set[@mv-object-type = '" + this.currentObjectType + "']") != null;
                    var productionHasImportFlows = this.ProductionXml.XPathSelectElement("//mv-data/import-attribute-flow/import-flow-set[@mv-object-type = '" + this.currentObjectType + "']") != null;
                    if (!pilotHasImportFlows && !productionHasImportFlows)
                    {
                        continue;
                    }

                    this.ProcessMetaverseObjectType();
                }

                // Sort by name
                var productionObjectTypes = from objectType in production
                                            let name = (string)objectType.Element(Documenter.DsmlNamespace + "name")
                                            orderby name
                                            select name;

                productionObjectTypes = productionObjectTypes.Where(productionObjectType => !pilotObjectTypes.Contains(productionObjectType));

                foreach (var objectType in productionObjectTypes)
                {
                    this.currentObjectType = objectType;
                    this.ProcessMetaverseObjectType();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current metaverse object type.
        /// </summary>
        private void ProcessMetaverseObjectType()
        {
            Logger.Instance.WriteMethodEntry("Metaverse Object Type: {0}.", this.currentObjectType);

            try
            {
                // Set Logger call context items
                Logger.SetContextItem(MetaverseDocumenter.LoggerContextItemMetaverseObjectType, this.currentObjectType);

                Logger.Instance.WriteInfo("Processing Metaverse Object Type.");

                this.CreateMetaverseObjectTypeDataSets();

                this.FillMetaverseObjectTypeDataSet(true);
                this.FillMetaverseObjectTypeDataSet(false);

                this.CreateMetaverseObjectTypeDiffgram();

                this.PrintMetaverseObjectType();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseObjectType);
            }
        }

        /// <summary>
        /// Creates the metaverse object type data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        private void CreateMetaverseObjectTypeDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("MetaverseObjectType") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Attribute");
                var column2 = new DataColumn("Type");
                var column3 = new DataColumn("Multi-valued");
                var column4 = new DataColumn("Indexed");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.Columns.Add(column4);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("MetaverseObjectTypePrecedence") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("Attribute");
                var column22 = new DataColumn("Precedence", typeof(int));
                var column32 = new DataColumn("PrecedenceDisplay"); // Rank or Manual or Equal
                var column42 = new DataColumn("Connector");
                var column52 = new DataColumn("SourceObjectType");
                var column62 = new DataColumn("SourceAttribute");
                var column72 = new DataColumn("MappingType");
                var column82 = new DataColumn("ConnectorGuid");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.Columns.Add(column42);
                table2.Columns.Add(column52);
                table2.Columns.Add(column62);
                table2.Columns.Add(column72);
                table2.Columns.Add(column82);
                table2.PrimaryKey = new[] { column12, column42, column52, column62, column72 };

                this.PilotDataSet = new DataSet("MetaverseObjectType") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetMetaverseObjectTypePrintTable();
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
        private DataTable GetMetaverseObjectTypePrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Multi-valued
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Indexed
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Precedence
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Precedence Display - Rank or Manual or Equal
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Connector
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 7 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Data Source Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Data Source Attribute
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 5 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Mapping Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 6 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // ConnectorGuid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 7 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the metaverse object type data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillMetaverseObjectTypeDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];
                var table2 = dataSet.Tables[1];

                var attributes = config.XPathSelectElements("//mv-data//dsml:class[dsml:name = '" + this.currentObjectType + "' ]/dsml:attribute", Documenter.NamespaceManager);

                // Sort by name
                attributes = from attribute in attributes
                             let name = (string)attribute.Attribute("ref")
                             orderby name
                             select attribute;

                for (var attributeIndex = 0; attributeIndex < attributes.Count(); ++attributeIndex)
                {
                    var attribute = attributes.ElementAt(attributeIndex);
                    var attributeName = ((string)attribute.Attribute("ref") ?? string.Empty).Trim('#');

                    // Set Logger call context items
                    Logger.SetContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute, attributeName);

                    Logger.Instance.WriteInfo("Processing Attribute Information.");

                    var attributeInfo = config.XPathSelectElement("//mv-data//dsml:attribute-type[dsml:name = '" + attributeName + "']", Documenter.NamespaceManager);

                    var attributeSyntax = (string)attributeInfo.Element(Documenter.DsmlNamespace + "syntax");

                    var row = table.NewRow();

                    row[0] = attributeName;
                    row[1] = Documenter.GetAttributeType(attributeSyntax, (string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexable"));
                    row[2] = ((string)attributeInfo.Attribute("single-value") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes";
                    row[3] = ((string)attributeInfo.Attribute(Documenter.MmsDsmlNamespace + "indexed") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No";

                    Documenter.AddRow(table, row);

                    Logger.Instance.WriteVerbose("Processed Attribute Information.");

                    // Fetch Import Flows
                    var importFlowsXPath = "//mv-data/import-attribute-flow/import-flow-set[@mv-object-type = '" + this.currentObjectType + "']/import-flows[@mv-attribute = '" + attributeName + "']";
                    var precedenceType = config.XPathSelectElement(importFlowsXPath) != null ? (string)config.XPathSelectElement(importFlowsXPath).Attribute("type") : string.Empty;
                    var importFlows = config.XPathSelectElements(importFlowsXPath + "/import-flow");
                    for (var importFlowIndex = 0; importFlowIndex < importFlows.Count(); ++importFlowIndex)
                    {
                        var importFlow = importFlows.ElementAt(importFlowIndex);
                        var row2 = table2.NewRow();

                        row2[0] = attributeName;

                        if (precedenceType.Equals("ranked", StringComparison.OrdinalIgnoreCase))
                        {
                            row2[1] = importFlowIndex + 1;
                            row2[2] = importFlowIndex + 1;
                        }
                        else
                        {
                            row2[1] = 1;
                            row2[2] = precedenceType;
                        }

                        var connectorId = ((string)importFlow.Attribute("src-ma") ?? string.Empty).ToUpperInvariant();
                        var connectorName = (string)config.XPathSelectElement("//ma-data[translate(id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connectorId + "']/name");
                        row2[3] = connectorName;

                        var connectorObjectType = (string)importFlow.Attribute("cd-object-type");
                        row2[4] = connectorObjectType;

                        Logger.Instance.WriteVerbose("Processing Sync Rule Info for Connector: '{0}'.", connectorName);

                        if (importFlow.XPathSelectElement("direct-mapping") != null)
                        {
                            row2[5] = (string)importFlow.XPathSelectElement("direct-mapping/src-attribute");
                            row2[6] = "Direct";
                        }
                        else if (importFlow.XPathSelectElement("scripted-mapping") != null)
                        {
                            var dataSourceAttribute = string.Empty;
                            foreach (var sourceAttribute in importFlow.XPathSelectElements("scripted-mapping/src-attribute"))
                            {
                                dataSourceAttribute += sourceAttribute + ",";
                            }

                            row2[5] = dataSourceAttribute.TrimEnd(',');
                            row2[6] = "Rules Extension - " + (string)importFlow.XPathSelectElement("scripted-mapping/script-context");
                        }
                        else if (importFlow.XPathSelectElement("constant-mapping") != null)
                        {
                            row2[5] = string.Empty;
                            row2[6] = "Constant - " + (string)importFlow.XPathSelectElement("constant-mapping/constant-value");
                        }
                        else if (importFlow.XPathSelectElement("dn-part--mapping") != null)
                        {
                            row2[5] = string.Empty;
                            row2[6] = "DN Component - ( " + (string)importFlow.XPathSelectElement("dn-part-mapping/dn-part") + " )";
                        }
                        else if (importFlow.XPathSelectElement("sync-rule-mapping") != null)
                        {
                            var mappingType = (string)importFlow.XPathSelectElement("sync-rule-mapping").Attribute("mapping-type") ?? string.Empty;
                            if (mappingType.Equals("direct", StringComparison.OrdinalIgnoreCase))
                            {
                                row2[5] = (string)importFlow.XPathSelectElement("sync-rule-mapping/src-attribute");
                                row2[6] = "Sync Rule - Direct";
                            }
                            else if (mappingType.Equals("constant", StringComparison.OrdinalIgnoreCase))
                            {
                                row2[5] = (string)importFlow.XPathSelectElement("sync-rule-mapping/sync-rule-value");
                                row2[6] = "Sync Rule - Constant";
                            }
                            else
                            {
                                var dataSourceAttribute = string.Empty;
                                foreach (var sourceAttribute in importFlow.XPathSelectElements("sync-rule-mapping/src-attribute"))
                                {
                                    dataSourceAttribute += sourceAttribute + ",";
                                }

                                row2[5] = dataSourceAttribute.TrimEnd(',');
                                row2[6] = "Sync Rule - Expression";  // TODO: Print the Sync Rule Expression
                            }
                        }

                        row2[7] = connectorId;
                        Documenter.AddRow(table2, row2);

                        Logger.Instance.WriteVerbose("Processed Sync Rule Info for Connector: '{0}'.", connectorName);
                    }
                }

                table.AcceptChanges();
                table2.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute);
            }
        }

        /// <summary>
        /// Creates the metaverse object type difference gram.
        /// </summary>
        private void CreateMetaverseObjectTypeDiffgram()
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
        /// Gets the metaverse object type header table.
        /// </summary>
        /// <returns>The metaverse object type header table.</returns>
        private DataTable GetMetaverseObjectTypeHeaderTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetHeaderTable();

                // Header Row 1
                // Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Attribute" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Type
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 1 }, { "ColumnName", "Type" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Multi-valued
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 2 }, { "ColumnName", "Multi-valued" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 5 } }).Values.Cast<object>().ToArray());

                // Indexed
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 3 }, { "ColumnName", "Indexed" }, { "RowSpan", 2 }, { "ColSpan", 1 }, { "ColWidth", 5 } }).Values.Cast<object>().ToArray());

                // Precedence
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 4 }, { "ColumnName", "Precedence" }, { "RowSpan", 1 }, { "ColSpan", 5 }, { "ColWidth", 0 } }).Values.Cast<object>().ToArray());

                // Header Row 2
                // Precedence Display - Rank or Manual or Equal
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 0 }, { "ColumnName", "Rank" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 5 } }).Values.Cast<object>().ToArray());

                // Connector
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 1 }, { "ColumnName", "Management Agent" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 15 } }).Values.Cast<object>().ToArray());

                // Data Source Object Type
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 2 }, { "ColumnName", "Data Source Object Type" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 10 } }).Values.Cast<object>().ToArray());

                // Data Source Attribute
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 3 }, { "ColumnName", "Data Source Attribute" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 20 } }).Values.Cast<object>().ToArray());

                // Mapping Type
                headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 1 }, { "ColumnIndex", 4 }, { "ColumnName", "Mapping Type" }, { "RowSpan", 1 }, { "ColSpan", 1 }, { "ColWidth", 20 } }).Values.Cast<object>().ToArray());

                headerTable.AcceptChanges();

                return headerTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the type of the metaverse object.
        /// </summary>
        private void PrintMetaverseObjectType()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = this.currentObjectType;

                this.WriteSectionHeader(sectionTitle, 4);

                var headerTable = this.GetMetaverseObjectTypeHeaderTable();
                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable, HtmlTableSize.Huge);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Metaverse ObjectType

        #region Metaverse Object Deletion Rules

        /// <summary>
        /// Processes the metaverse object deletion rules.
        /// </summary>
        private void ProcessMetaverseObjectDeletionRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Processing Metaverse Object Deletion Rules Summary");

                const string XPath = "//mv-data//dsml:class";
                var pilot = this.PilotXml.XPathSelectElements(XPath, Documenter.NamespaceManager);
                var production = this.ProductionXml.XPathSelectElements(XPath, Documenter.NamespaceManager);

                // Sort by name
                var pilotObjectTypes = from objectType in pilot
                                       let name = (string)objectType.Element(Documenter.DsmlNamespace + "name")
                                       orderby name
                                       select name;

                foreach (var objectType in pilotObjectTypes)
                {
                    this.ProcessMetaverseObjectDeletionRule(objectType);
                }

                // Sort by name
                var productionObjectTypes = from objectType in production
                                            let name = (string)objectType.Element(Documenter.DsmlNamespace + "name")
                                            orderby name
                                            select name;

                productionObjectTypes = productionObjectTypes.Where(productionObjectType => !pilotObjectTypes.Contains(productionObjectType));

                foreach (var objectType in productionObjectTypes)
                {
                    this.ProcessMetaverseObjectDeletionRule(objectType);
                }

                this.PrintMetaverseObjectDeletionRules();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the metaverse object deletion rule.
        /// </summary>
        /// <param name="objectType">Name of the object.</param>
        private void ProcessMetaverseObjectDeletionRule(string objectType)
        {
            Logger.Instance.WriteMethodEntry("Metaverse Object Type: '{0}'.", objectType);

            try
            {
                // Set Logger call context items
                Logger.SetContextItem(MetaverseDocumenter.LoggerContextItemMetaverseObjectType, objectType);

                Logger.Instance.WriteInfo("Processing Metaverse Object Deletion Rules.");

                this.CreateMetaverseObjectDeletionRuleDataSets();

                this.FillMetaverseObjectDeletionRuleDataSet(objectType, true);
                this.FillMetaverseObjectDeletionRuleDataSet(objectType, false);

                this.CreateMetaverseObjectDeletionRuleDiffgram();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseObjectType);
            }
        }

        /// <summary>
        /// Creates the metaverse object deletion rule data sets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        private void CreateMetaverseObjectDeletionRuleDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("MetaverseObjectTypes") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Object Type");
                var column2 = new DataColumn("Deletion Rule Type");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1, column2 };

                var table2 = new DataTable("MetaverseObjectTypeConnectors") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("Object Type");
                var column22 = new DataColumn("Connector");
                var column32 = new DataColumn("ConnectorGuid");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.PrimaryKey = new[] { column12, column22 };

                this.PilotDataSet = new DataSet("MetaverseObjectDeletionRules") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);

                this.ProductionDataSet = this.PilotDataSet.Clone();

                var printTable = this.GetMetaverseObjectDeletionRulePrintTable();
                this.PilotDataSet.Tables.Add(printTable);
                this.ProductionDataSet.Tables.Add(printTable.Copy());
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the metaverse object deletion rule print table.
        /// </summary>
        /// <returns>The metaverse object deletion rule print table.</returns>
        private DataTable GetMetaverseObjectDeletionRulePrintTable()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 1
                // Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Deletion Rule Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                // Object Type
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 0 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Connector
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", 2 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // ConnectorGuid
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());

                printTable.AcceptChanges();

                return printTable;
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the metaverse object deletion rule data set.
        /// </summary>
        /// <param name="objectType">Type of the object.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillMetaverseObjectDeletionRuleDataSet(string objectType, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];
                var table2 = dataSet.Tables[1];

                var deletionRule = config.XPathSelectElement("//mv-data/mv-deletion/mv-deletion-rule[@mv-object-type = '" + objectType + "']");

                if (deletionRule != null)
                {
                    var deletionRuleType = (string)deletionRule.Attribute("type") ?? string.Empty;

                    if (deletionRuleType.Equals("scripted", StringComparison.OrdinalIgnoreCase))
                    {
                        deletionRuleType = "Rules Extension";
                    }
                    else if (deletionRuleType.Equals("declared-any", StringComparison.OrdinalIgnoreCase))
                    {
                        deletionRuleType = "Delete the metaverse object when connector from any of the following management agents is disconnected";
                    }
                    else if (deletionRuleType.Equals("declared-last", StringComparison.OrdinalIgnoreCase))
                    {
                        deletionRuleType = "Delete the metaverse object when the last connector is disconnected. Ignore from the following list of management agents";
                    }

                    Documenter.AddRow(table, new object[] { objectType, deletionRuleType });

                    foreach (var connector in deletionRule.Elements("src-ma"))
                    {
                        var connectorId = ((string)connector ?? string.Empty).ToUpperInvariant();
                        var connectorName = (string)config.XPathSelectElement("//ma-data[translate(id, '" + Documenter.LowercaseLetters + "', '" + Documenter.UppercaseLetters + "') = '" + connectorId + "']/name");

                        var row2 = table2.NewRow();

                        row2[0] = objectType;
                        row2[1] = connectorName;
                        row2[2] = connectorId;

                        Documenter.AddRow(table2, row2);
                    }
                }

                table.AcceptChanges();
                table2.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Creates the metaverse object deletion rule difference gram.
        /// </summary>
        private void CreateMetaverseObjectDeletionRuleDiffgram()
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
        /// Prints the metaverse object deletion rule.
        /// </summary>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        private void PrintMetaverseObjectDeletionRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Metaverse Object Deletion Rules";
                this.WriteSectionHeader(sectionTitle, 3);

                if (this.DiffgramDataSets.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Object Type", 15 }, { "Deletion Rule Type", 60 }, { "Management Agents", 35 } });

                    #region table

                    this.ReportWriter.WriteBeginTag("table");
                    this.ReportWriter.WriteAttribute("class", HtmlTableSize.Standard.ToString() + " " + this.GetCssVisibilityClass());
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        #region thead

                        this.WriteTableHeader(headerTable);

                        #endregion thead
                    }

                    #region rows

                    foreach (var dataSet in this.DiffgramDataSets)
                    {
                        this.DiffgramDataSet = dataSet;
                        this.WriteRows(dataSet.Tables[0].Rows);
                    }

                    #endregion rows

                    this.ReportWriter.WriteEndTag("table");

                    #endregion table
                }
                else
                {
                    this.WriteContentParagraph("There are no metaverse object deletion rules configured.");
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Metaverse Object Deletion Rules

        #region Metaverse Options

        /// <summary>
        /// Processes the metaverse options.
        /// </summary>
        private void ProcessMetaverseOptions()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Metaverse Options");

                var sectionTitle = "Metaverse Options";
                this.WriteSectionHeader(sectionTitle, 3);

                this.ProcessMetaverseRulesExtension();
                this.ProcessSynchronizationRuleSettings();
                this.ProcessGlobalRulesExtensionSettings();
                this.ProcessWMIPasswordManagementSettings();
                this.ProcessPasswordSynchronization();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Metaverse Rules Extension

        /// <summary>
        /// Processes the metaverse rules extension.
        /// </summary>
        private void ProcessMetaverseRulesExtension()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Metaverse Rules Extension");

                this.CreateSimpleSettingsDataSets(2);  // 1 = Name, 2 = Value

                this.FillMetaverseRulesExtensionDataSet(true);
                this.FillMetaverseRulesExtensionDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintMetaverseRulesExtension();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the metaverse rules extension data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillMetaverseRulesExtensionDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                var assemblyName = (string)config.XPathSelectElement("//mv-data//extension/assembly-name") ?? string.Empty;

                if (string.IsNullOrEmpty(assemblyName))
                {
                    Documenter.AddRow(table, new object[] { "Enable metaverse rules extension", "No" });
                }
                else
                {
                    Documenter.AddRow(table, new object[] { "Enable metaverse rules extension", "Yes" });
                    Documenter.AddRow(table, new object[] { "Rules extension name", assemblyName });
                    Documenter.AddRow(table, new object[] { "Run this rules extension in a separate process", config.XPathSelectElement("//mv-data//extension[application-protection = 'high']") != null ? "Yes" : "No" });
                    Documenter.AddRow(table, new object[] { "Enable Provisioning Rules Extension", config.XPathSelectElement("//mv-data//provisioning[@type = 'both' or @type = 'scripted']") != null ? "Yes" : "No" });
                }

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute);
            }
        }

        /// <summary>
        /// Prints the metaverse rules extension.
        /// </summary>
        private void PrintMetaverseRulesExtension()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Metaverse Rules Extension", 60 }, { string.Empty, 40 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Metaverse Rules Extension

        #region Synchronization Rule Settings

        /// <summary>
        /// Processes the synchronization rule settings.
        /// </summary>
        private void ProcessSynchronizationRuleSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Synchronization Rule Settings");

                this.CreateSimpleSettingsDataSets(2);  // 1 = Name, 2 = Value

                this.FillSynchronizationRuleSettingsDataSet(true);
                this.FillSynchronizationRuleSettingsDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintSynchronizationRuleSettings();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the synchronization rule settings data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillSynchronizationRuleSettingsDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                Documenter.AddRow(table, new object[] { "Enable Synchronization Rules Provisioning", config.XPathSelectElement("//mv-data//provisioning[@type = 'both' or @type = 'sync-rule']") != null ? "Yes" : "No" });

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute);
            }
        }

        /// <summary>
        /// Prints the synchronization rule settings.
        /// </summary>
        private void PrintSynchronizationRuleSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Synchronization Rule Settings", 60 }, { string.Empty, 40 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Synchronization Rule Settings

        #region Global Rules Extension Settings

        /// <summary>
        /// Processes the global rules extension settings.
        /// </summary>
        private void ProcessGlobalRulesExtensionSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Global Rules Extension Settings");

                this.CreateSimpleSettingsDataSets(2);  // 1 = Name, 2 = Value

                this.FillGlobalRulesExtensionSettingsDataSet(true);
                this.FillGlobalRulesExtensionSettingsDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintGlobalRulesExtensionSettings();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the global rules extension settings data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillGlobalRulesExtensionSettingsDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                Documenter.AddRow(table, new object[] { "Unload extension if the duration of single operation exceeds (seconds)", (string)config.XPathSelectElement("//mv-data//extension/timeout") });

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute);
            }
        }

        /// <summary>
        /// Prints the global rules extension settings.
        /// </summary>
        private void PrintGlobalRulesExtensionSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Global Rules Extension Settings", 60 }, { string.Empty, 40 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Global Rules Extension Settings

        #region WMI Password Management Settings

        /// <summary>
        /// Processes the WMI password management settings.
        /// </summary>
        private void ProcessWMIPasswordManagementSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("WMI Password Management Settings");

                this.CreateSimpleSettingsDataSets(2);  // 1 = Name, 2 = Value

                this.FillWMIPasswordManagementSettingsDataSet(true);
                this.FillWMIPasswordManagementSettingsDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintWMIPasswordManagementSettings();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the WMI password management settings data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillWMIPasswordManagementSettingsDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                Documenter.AddRow(table, new object[] { "Save last x password change/set event details", (string)config.XPathSelectElement("//mv-data/password-change-history-size") });

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute);
            }
        }

        /// <summary>
        /// Prints the WMI password management settings.
        /// </summary>
        private void PrintWMIPasswordManagementSettings()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "WMI Password Management Settings", 60 }, { string.Empty, 40 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion WMI Password Management Settings

        #region Password Synchronization

        /// <summary>
        /// Processes the password synchronization.
        /// </summary>
        private void ProcessPasswordSynchronization()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                Logger.Instance.WriteInfo("Password Synchronization");

                this.CreateSimpleSettingsDataSets(2);  // 1 = Name, 2 = Value

                this.FillPasswordSynchronizationDataSet(true);
                this.FillPasswordSynchronizationDataSet(false);

                this.CreateSimpleSettingsDiffgram();

                this.PrintPasswordSynchronization();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the password synchronization data set.
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        private void FillPasswordSynchronizationDataSet(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var config = pilotConfig ? this.PilotXml : this.ProductionXml;
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;

                var table = dataSet.Tables[0];

                Documenter.AddRow(table, new object[] { "Enable Password Synchronization", config.XPathSelectElement("//mv-data/password-sync[password-sync-enabled = '1']") != null ? "Yes" : "No" });

                table.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);

                // Clear Logger call context items
                Logger.ClearContextItem(MetaverseDocumenter.LoggerContextItemMetaverseAttribute);
            }
        }

        /// <summary>
        /// Prints the password synchronization.
        /// </summary>
        private void PrintPasswordSynchronization()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Password Synchronization", 60 }, { string.Empty, 40 } });

                this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Password Synchronization

        #endregion Metaverse Options
    }
}
