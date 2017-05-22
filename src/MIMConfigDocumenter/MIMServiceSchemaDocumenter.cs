//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="MIMServiceSchemaDocumenter.cs" company="Microsoft">
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
    /// The MIMServiceSchemaDocumenter documents the configuration of an MIM Service deployment.
    /// </summary>
    public class MIMServiceSchemaDocumenter : ServiceCommonDocumenter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MIMServiceSchemaDocumenter"/> class.
        /// </summary>
        /// <param name="pilotConfigXml">The pilot configuration XML.</param>
        /// <param name="productionConfigXml">The production configuration XML.</param>
        /// <param name="changesConfigXml">The changes configuration XML.</param>
        public MIMServiceSchemaDocumenter(XElement pilotConfigXml, XElement productionConfigXml, XElement changesConfigXml)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PilotXml = pilotConfigXml;
                this.ProductionXml = productionConfigXml;
                this.ChangesXml = changesConfigXml;

                this.ReportFileName = Documenter.GetTempFilePath("Schema.tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath("Schema.TOC.tmp.html");
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the MIM Service schema configuration report.
        /// </summary>
        /// <returns>
        /// The Tuple of configuration report and associated TOC
        /// </returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed. XhtmlTextWriter takes care of disposting StreamWriter.")]
        public override Tuple<string, string> GetReport()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.ReportWriter = new XhtmlTextWriter(new StreamWriter(this.ReportFileName));
                this.ReportToCWriter = new XhtmlTextWriter(new StreamWriter(this.ReportToCFileName));

                var sectionTitle = "Schema Customizations";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 2);

                this.ProcessObjectTypeDescriptions();
                this.ProcessAttributesAndBindingSummary();
                this.ProcessAttributeTypeDescriptions();
                this.ProcessBindingDescriptions();

                return this.GetReportTuple();
            }
            catch (Exception e)
            {
                throw Logger.Instance.ReportError(e);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region ObjectTypeDescription Configuration

        /// <summary>
        /// Processes the ObjectTypeDescription objects.
        /// </summary>
        protected void ProcessObjectTypeDescriptions()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Resource Types";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "ObjectTypeDescription";

                var changeObjects = this.GetChangedObjects(objectType);

                if (changeObjects.Count() == 0)
                {
                    this.WriteContentParagraph(string.Format(CultureInfo.CurrentUICulture, DocumenterResources.NoCustomizationsDetected, ServiceCommonDocumenter.GetObjectTypeDisplayName(objectType)));
                }
                else
                {
                    foreach (var changeObject in changeObjects)
                    {
                        this.CurrentChangeObject = changeObject;
                        this.ProcessObjectTypeDescription();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current ObjectTypeDescription  object.
        /// </summary>
        protected void ProcessObjectTypeDescription()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PrintSimpleSectionHeader(4);

                // General Info
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Resource Type", "ObjectType"),
                    new KeyValuePair<string, string>("System Name", "Name"),
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                // Localization Info
                this.ProcessLocalizationConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the current ObjectTypeDescription object section header.
        /// </summary>
        protected void PrintObjectTypeDescriptionSectionHeader()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = this.GetAttributeChange("DisplayName").NewValue;
                var sectionGuid = this.GetAttributeChange("ObjectID").NewValue;

                Logger.Instance.WriteInfo("Processing ObjectTypeDescription:  " + sectionTitle);

                this.WriteSectionHeader(sectionTitle, 4, sectionGuid);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion ObjectTypeDescription Configuration

        #region Attributes And Bindings Summary

        /// <summary>
        /// Processes the attributes and bindings summary.
        /// </summary>
        protected void ProcessAttributesAndBindingSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Attributes and Bindings Summary";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                this.ProcessAttributesSummary();
                this.ProcessBindingsSummary();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region AttributeTypeDescription Summary

        /// <summary>
        /// Processes the attributes summary.
        /// </summary>
        protected void ProcessAttributesSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Attributes Summary";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 4, ConfigEnvironment.PilotAndProduction);

                this.CreateAttributeTypeDescriptionSummaryDiffgramDataSet();
                this.FillAttributeTypeDescriptionSummaryDiffgramDataSet();
                this.PrintAttributeTypeDescriptionSummary();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates AttributeTypeDescription summary diffgram dataset
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateAttributeTypeDescriptionSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("AttributeTypeDescriptionSummary") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Display Name");
                var column2 = new DataColumn("Display Name Markup");
                var column3 = new DataColumn("System Name");
                var column4 = new DataColumn("Data Type");
                var column5 = new DataColumn("Description");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.Columns.Add(column4);
                table.Columns.Add(column5);
                table.PrimaryKey = new[] { column3, column4 };

                var printTable = Documenter.GetPrintTable();

                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                var diffgramTable = Documenter.CreateDiffgramTable(table);

                this.DiffgramDataSet = new DataSet("AttributeTypeDescriptionSummary") { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSet.Tables.Add(diffgramTable);
                this.DiffgramDataSet.Tables.Add(printTable);

                Documenter.AddRowVisibilityStatusColumn(this.DiffgramDataSet);

                this.DiffgramDataSet.AcceptChanges();
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the AttributeTypeDescription summary dataset.
        /// </summary>
        protected void FillAttributeTypeDescriptionSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];

                var objectType = "AttributeTypeDescription";

                var changeObjects = this.GetChangedObjects(objectType);

                if (changeObjects.Count() == 0)
                {
                    this.WriteContentParagraph(string.Format(CultureInfo.CurrentUICulture, DocumenterResources.NoCustomizationsDetected, ServiceCommonDocumenter.GetObjectTypeDisplayName(objectType)));
                }
                else
                {
                    foreach (var changeObject in changeObjects)
                    {
                        this.CurrentChangeObject = changeObject;
                        var state = this.CurrentChangeObjectState;
                        var objectModificationType = state == "Create" ? DataRowState.Added : state == "Delete" ? DataRowState.Deleted : DataRowState.Modified;
                        var attributeObjectId = this.GetAttributeChange("ObjectID");
                        var displayNameChange = this.GetAttributeChange("DisplayName");
                        var systemNameChange = this.GetAttributeChange("Name");
                        var dataTypeChange = this.GetAttributeChange("DataType");
                        var descriptionChange = this.GetAttributeChange("Description");

                        var objectId = attributeObjectId.NewValue;
                        var objectIdOld = attributeObjectId.OldValue;
                        var displayName = displayNameChange.NewValue;
                        var displayNameOld = displayNameChange.OldValue;
                        var displayNameMarkupNew = ServiceCommonDocumenter.GetJumpToBookmarkLocationMarkup(displayName, objectId, objectModificationType);
                        var displayNameMarkupOld = ServiceCommonDocumenter.GetJumpToBookmarkLocationMarkup(displayNameOld, objectIdOld, objectModificationType);
                        var systemName = systemNameChange.NewValue;
                        var dataType = dataTypeChange.NewValue;
                        var description = descriptionChange.NewValue;
                        var descriptionOld = descriptionChange.OldValue;

                        Documenter.AddRow(diffgramTable, new object[] { displayName, displayNameMarkupNew, systemName, dataType, description, objectModificationType, displayNameOld, displayNameMarkupOld, descriptionOld });
                    }

                    this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the AttributeTypeDescription summary.
        /// </summary>
        protected void PrintAttributeTypeDescriptionSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Display Name", 35 }, { "System Name", 25 }, { "Data Type", 10 }, { "Description", 30 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion AttributeTypeDescription Summary

        #region BindingDescription Summary

        /// <summary>
        /// Processes the bindings summary.
        /// </summary>
        protected void ProcessBindingsSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Bindings Summary";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 4, ConfigEnvironment.PilotAndProduction);

                this.CreateBindingDescriptionSummaryDiffgramDataSet();
                this.FillBindingDescriptionSummaryDiffgramDataSet();
                this.PrintBindingDescriptionSummary();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the BindingDescription summary diffgram dataset.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateBindingDescriptionSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("BindingDescriptionSummary") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Bound Object Type");
                var column2 = new DataColumn("Bound Object Type Markup");
                var column3 = new DataColumn("Bound Attribute Type");
                var column4 = new DataColumn("Bound Attribute Type Markup");
                var column5 = new DataColumn("Required");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.Columns.Add(column4);
                table.Columns.Add(column5);
                table.PrimaryKey = new[] { column1, column3 };

                var printTable = Documenter.GetPrintTable();

                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                var diffgramTable = Documenter.CreateDiffgramTable(table);

                this.DiffgramDataSet = new DataSet("BindingDescriptionSummary") { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSet.Tables.Add(diffgramTable);
                this.DiffgramDataSet.Tables.Add(printTable);

                Documenter.AddRowVisibilityStatusColumn(this.DiffgramDataSet);

                this.DiffgramDataSet.AcceptChanges();
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the BindingDescription summary dataset.
        /// </summary>
        protected void FillBindingDescriptionSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];

                var objectType = "BindingDescription";

                var changeObjects = this.GetChangedObjects(objectType);

                if (changeObjects.Count() == 0)
                {
                    this.WriteContentParagraph(string.Format(CultureInfo.CurrentUICulture, DocumenterResources.NoCustomizationsDetected, ServiceCommonDocumenter.GetObjectTypeDisplayName(objectType)));
                }
                else
                {
                    foreach (var changeObject in changeObjects)
                    {
                        this.CurrentChangeObject = changeObject;
                        var state = this.CurrentChangeObjectState;
                        var objectModificationType = state == "Create" ? DataRowState.Added : state == "Delete" ? DataRowState.Deleted : DataRowState.Modified;
                        var bindingDescriptionIdChange = this.GetAttributeChange("ObjectID");
                        var boundObjectTypeChange = this.GetAttributeChange("BoundObjectType");
                        var boundAttributeTypeChange = this.GetAttributeChange("BoundAttributeType");
                        var requiredChange = this.GetAttributeChange("Required");

                        var boundObjectType = boundObjectTypeChange.NewValueText;
                        var boundObjectTypeMarkup = ServiceCommonDocumenter.GetJumpToBookmarkLocationMarkup(boundObjectType, bindingDescriptionIdChange.NewValue, objectModificationType);
                        var boundObjectTypeMarkupOld = boundObjectTypeMarkup;
                        var boundAttributeType = boundAttributeTypeChange.NewValueText;
                        var boundAttributeTypeId = boundAttributeTypeChange.NewId;
                        var boundAttributeTypeMarkup = ServiceCommonDocumenter.GetJumpToBookmarkLocationMarkup(boundAttributeType, boundAttributeTypeId, objectModificationType);
                        var boundAttributeTypeMarkupOld = boundAttributeTypeMarkup;
                        var required = requiredChange.NewValue;
                        var oldRequired = requiredChange.OldValue;

                        Documenter.AddRow(diffgramTable, new object[] { boundObjectType, boundObjectTypeMarkup, boundAttributeType, boundAttributeTypeMarkup, required, objectModificationType, boundObjectTypeMarkupOld, boundAttributeTypeMarkupOld, oldRequired });
                    }

                    this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the BindingDescription summary.
        /// </summary>
        protected void PrintBindingDescriptionSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Bound Object Type", 40 }, { "Bound Attribute Type", 40 }, { "Required", 20 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion BindingDescription Summary

        #endregion Attributes And Bindings Summary

        #region AttributeTypeDescription Configuration

        /// <summary>
        /// Processes the AttributeTypeDescription objects.
        /// </summary>
        protected void ProcessAttributeTypeDescriptions()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Attributes";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "AttributeTypeDescription";

                var changeObjects = this.GetChangedObjects(objectType);

                if (changeObjects.Count() == 0)
                {
                    this.WriteContentParagraph(string.Format(CultureInfo.CurrentUICulture, DocumenterResources.NoCustomizationsDetected, ServiceCommonDocumenter.GetObjectTypeDisplayName(objectType)));
                }
                else
                {
                    foreach (var changeObject in changeObjects)
                    {
                        this.CurrentChangeObject = changeObject;
                        this.ProcessAttributeTypeDescription();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current AttributeTypeDescription object.
        /// </summary>
        protected void ProcessAttributeTypeDescription()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PrintSimpleSectionHeader(4);

                // General Info
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Resource Type", "ObjectType"),
                    new KeyValuePair<string, string>("System Name", "Name"),
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Data Type", "DataType"),
                    new KeyValuePair<string, string>("Multivalued", "Multivalued"),
                    new KeyValuePair<string, string>("Description", "Description"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                // Localization Info
                this.ProcessLocalizationConfiguration();

                // Validation Configuration
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("String pattern", "StringRegex"),
                    new KeyValuePair<string, string>("Minimum Inclusive Integer", "IntegerMinimum"),
                    new KeyValuePair<string, string>("Maximum Inclusive Integer", "IntegerMaximum"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Validation", 30 }, { string.Empty, 70 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion AttributeTypeDescription Configuration

        #region BindingDescription Configuration

        /// <summary>
        /// Processes the BindingDescription objects.
        /// </summary>
        protected void ProcessBindingDescriptions()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Bindings";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "BindingDescription";

                var changeObjects = this.GetChangedObjects(objectType);

                if (changeObjects.Count() == 0)
                {
                    this.WriteContentParagraph(string.Format(CultureInfo.CurrentUICulture, DocumenterResources.NoCustomizationsDetected, ServiceCommonDocumenter.GetObjectTypeDisplayName(objectType)));
                }
                else
                {
                    foreach (var changeObject in changeObjects)
                    {
                        this.CurrentChangeObject = changeObject;
                        this.ProcessBindingDescription();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current BindingDescription object.
        /// </summary>
        protected void ProcessBindingDescription()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PrintBindingDescriptionSectionHeader();

                // General Info
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Resource Type", "BoundObjectType"),
                    new KeyValuePair<string, string>("Attribute Type", "BoundAttributeType"),
                    new KeyValuePair<string, string>("Required", "Required"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                // Attribute Override Info
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Attribute Override", 30 }, { string.Empty, 70 } });

                // Localization Info
                this.ProcessLocalizationConfiguration();

                // Validation Configuration
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("String pattern", "StringRegex"),
                    new KeyValuePair<string, string>("Minimum Inclusive Integer", "IntegerMinimum"),
                    new KeyValuePair<string, string>("Maximum Inclusive Integer", "IntegerMaximum"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Validation", 30 }, { string.Empty, 70 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the current BindingDescription object.
        /// </summary>
        protected void PrintBindingDescriptionSectionHeader()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var boundObjectTypeChange = this.GetAttributeChange("BoundObjectType");
                var boundAttributeTypeChange = this.GetAttributeChange("BoundAttributeType");
                var bookmark = boundObjectTypeChange.NewValueText;
                var sectionTitle = bookmark + " - " + boundAttributeTypeChange.NewValueText;
                var sectionGuid = this.GetAttributeChange("ObjectID").NewValue;

                Logger.Instance.WriteInfo("Processing " + this.CurrentChangeObjectType + ":  " + sectionTitle);

                this.WriteSectionHeader(sectionTitle, 4, bookmark, sectionGuid);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion BindingDescription Configuration
    }
}
