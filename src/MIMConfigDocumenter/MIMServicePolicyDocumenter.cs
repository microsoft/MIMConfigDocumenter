//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="MIMServicePolicyDocumenter.cs" company="Microsoft">
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
    using System.Collections;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Web.UI;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The MIMServicePolicyDocumenter documents the configuration of an MIM Service deployment.
    /// </summary>
    public class MIMServicePolicyDocumenter : ServiceCommonDocumenter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MIMServicePolicyDocumenter"/> class.
        /// </summary>
        /// <param name="pilotConfigXml">The pilot configuration XML.</param>
        /// <param name="productionConfigXml">The production configuration XML.</param>
        /// <param name="changesConfigXml">The changes configuration XML.</param>
        public MIMServicePolicyDocumenter(XElement pilotConfigXml, XElement productionConfigXml, XElement changesConfigXml)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PilotXml = pilotConfigXml;
                this.ProductionXml = productionConfigXml;
                this.ChangesXml = changesConfigXml;

                this.ReportFileName = Documenter.GetTempFilePath("Policy.tmp.html");
                this.ReportToCFileName = Documenter.GetTempFilePath("Policy.TOC.tmp.html");
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the MIM Service policy configuration report.
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

                var sectionTitle = "Policy Customizations";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 2);

                this.ProcessActivityInformationConfigurations();
                this.ProcessForestAndDomainConfigurations();
                this.ProcessFilterPermissions();
                this.ProcessSetsSummary();
                this.ProcessSets();
                this.ProcessManagementPolicyRulesSummary();
                this.ProcessManagementPolicyRules();
                this.ProcessWorkflowsSummary();
                this.ProcessWorkflows();
                this.ProcessSynchronizationRulesSummary();
                this.ProcessSynchronizationRules();
                this.ProcessEmailTemplates();

                // Portal UI Configurations
                sectionTitle = "Portal Configuration";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 2);
                this.ProcessObjectVisualizationConfigurations();
                this.ProcessPortalUIConfigurations();
                this.ProcessHomepageConfigurations();
                this.ProcessNavigationBarConfigurations();
                this.ProcessSearchScopeConfigurations();

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

        #region ActivityInformationConfiguration Configurations

        /// <summary>
        /// Processes the ActivityInformationConfiguration objects.
        /// </summary>
        protected void ProcessActivityInformationConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Activity Information Configurations";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "ActivityInformationConfiguration";

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
                        this.ProcessActivityInformationConfiguration();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current ActivityInformationConfiguration  object.
        /// </summary>
        protected void ProcessActivityInformationConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PrintSimpleSectionHeader(4);

                // Common Attributes
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Resource Type", "ObjectType"),
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Common Attributes", 30 }, { string.Empty, 70 } });

                // Extended Attributes
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Activity Name", "ActivityName"),
                    new KeyValuePair<string, string>("Assembly Name", "AssemblyName"),
                    new KeyValuePair<string, string>("Is Action Activity", "IsActionActivity"),
                    new KeyValuePair<string, string>("Is Authentication Activity", "IsAuthenticationActivity"),
                    new KeyValuePair<string, string>("Is Authorization Activity", "IsAuthorizationActivity"),
                    new KeyValuePair<string, string>("Type Name", "TypeName"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Extended Attributes", 30 }, { string.Empty, 70 } });

                // Localization Info
                this.ProcessLocalizationConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion ActivityInformationConfiguration Configurations

        #region Forest and Domian Configurations

        /// <summary>
        /// Processes the forest and domain configuration objects.
        /// </summary>
        protected void ProcessForestAndDomainConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Forest and Domain Configurations";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                this.ProcessForestConfigurations();
                this.ProcessDomainConfigurations();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Forest Configurations

        /// <summary>
        /// Processes the forest configuration objects.
        /// </summary>
        protected void ProcessForestConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Forest Configurations";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 4, ConfigEnvironment.PilotAndProduction);

                var objectType = "ForestConfiguration";

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
                        this.ProcessForestConfiguration();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current ForestConfiguration  object.
        /// </summary>
        protected void ProcessForestConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PrintSimpleSectionHeader(5);

                // Basics
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Resource Type", "ObjectType"),
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Trusted Forest", "TrustedForest"),
                    new KeyValuePair<string, string>("Distribution Group Domain", "DistributionListDomain"),
                    new KeyValuePair<string, string>("Contact Set", "ContactSet"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Basics", 30 }, { string.Empty, 70 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Forest Configurations

        #region Domain Configurations

        /// <summary>
        /// Processes the domain configuration objects.
        /// </summary>
        protected void ProcessDomainConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Domain Configurations";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 4, ConfigEnvironment.PilotAndProduction);

                var objectType = "DomainConfiguration";

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
                        this.ProcessDomainConfiguration();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current DomainConfiguration  object.
        /// </summary>
        protected void ProcessDomainConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PrintSimpleSectionHeader(5);

                // General
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Resource Type", "ObjectType"),
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Domain", "Domain"),
                    new KeyValuePair<string, string>("Forest Configuration", "ForestConfiguration"),
                    new KeyValuePair<string, string>("Foreign Security Principal Set", "ForeignSecurityPrincipalSet"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Domain Configurations

        #endregion Forest and Domian Configurations

        #region Filter Permissions

        /// <summary>
        /// Processes the FilterScope objects.
        /// </summary>
        protected void ProcessFilterPermissions()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Filter Permissions";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "FilterScope";

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
                        this.ProcessFilterPermission();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current FilterScope  object.
        /// </summary>
        protected void ProcessFilterPermission()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PrintSimpleSectionHeader(4);

                // General
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Resource Type", "ObjectType"),
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Allowed Attributes", "AllowedAttributes"),
                    new KeyValuePair<string, string>("Allowed Membership References", "AllowedMembershipReferences"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Filter Permissions

        #region Sets

        /// <summary>
        /// Gets the value of filter string
        /// </summary>
        /// <param name="filter">The set filter xml string</param>
        /// <returns>The value of filter string</returns>
        protected string GetFilterText(string filter)
        {
            if (!string.IsNullOrEmpty(filter))
            {
                try
                {
                    var xmlFilter = XElement.Parse(filter);
                    filter = xmlFilter.Value;
                }
                catch (Exception e)
                {
                    var info = string.Format(CultureInfo.InvariantCulture, "Unable to Parse Filter '{0}'. Error {1}.", filter, e.ToString());
                    Logger.Instance.WriteInfo(info);
                }
            }

            return filter;
        }

        #region Sets Summary

        /// <summary>
        /// Processes the sets summary.
        /// </summary>
        protected void ProcessSetsSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Sets Summary";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                this.CreateSetsSummaryDiffgramDataSet();
                this.FillSetsSummaryDiffgramDataSet();
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Display Name", 30 }, { "Criteria-Based Membership Filter", 70 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the Set summary diffgram dataset.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateSetsSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("SetsSummary") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Display Name"); // for sorting purposes
                var column2 = new DataColumn("Display Name Markup");
                var column3 = new DataColumn("Filter Markup");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.PrimaryKey = new[] { column1 };

                var printTable = Documenter.GetPrintTable();

                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                var diffgramTable = Documenter.CreateDiffgramTable(table);

                this.DiffgramDataSet = new DataSet("SetsSummary") { Locale = CultureInfo.InvariantCulture };
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
        /// Fills the sets summary diffgram dataset
        /// </summary>
        protected void FillSetsSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];

                var objectType = "Set";

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
                        var attributeObjectIdChange = this.GetAttributeChange("ObjectID");
                        var displayNameChange = this.GetAttributeChange("DisplayName");
                        var filterChange = this.GetAttributeChange("Filter");

                        var objectId = attributeObjectIdChange.NewValue;
                        var displayName = displayNameChange.NewValue;
                        var displayNameMarkup = ServiceCommonDocumenter.GetJumpToBookmarkLocationMarkup(displayName, objectId, objectModificationType);
                        var filter = this.GetFilterText(filterChange.NewValue);
                        var filterOld = this.GetFilterText(filterChange.OldValue);

                        Documenter.AddRow(diffgramTable, new object[] { displayName, displayNameMarkup, filter, objectModificationType, displayNameMarkup, filterOld });
                    }

                    this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Sets Summary

        #region Sets Details

        /// <summary>
        /// Processes the sets detailed configuration.
        /// </summary>
        protected void ProcessSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Sets";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "Set";

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
                        this.ProcessSet();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current Set object.
        /// </summary>
        protected void ProcessSet()
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
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Manually-managed Members", "ExplicitMember"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                // Criteria-based Members
                this.CreateSimpleMultivalueValuesDiffgramDataSet();

                var filterChange = this.GetAttributeChange("Filter");
                var filter = this.GetFilterText(filterChange.NewValue);
                var filterOld = this.GetFilterText(filterChange.OldValue);

                var diffgramTable = this.DiffgramDataSet.Tables[0];
                Documenter.AddRow(diffgramTable, new object[] { filter, filter, filterChange.AttributeModificationType, filterOld, filterOld });

                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Criteria-based Members", 100 } });

                // Manually-managed Members
                this.CreateSimpleMultivalueValuesDiffgramDataSet();
                this.FillSimpleMultivalueValuesDiffgramDataSet("ExplicitMember");
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Manually-managed Members", 100 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Sets Details

        #endregion Sets

        #region MPRs

        #region MPRs Summary

        /// <summary>
        /// Processes the MPR summary.
        /// </summary>
        protected void ProcessManagementPolicyRulesSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Management Policy Rules Summary";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                this.CreateManagementPolicyRulesSummaryDiffgramDataSet();
                this.FillManagementPolicyRulesSummaryDiffgramDataSet();
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Display Name", 40 }, { "Disabled", 10 }, { "Grant Permission", 10 }, { "AuthN Workflows", 10 }, { "AuthZ Workflows", 10 }, { "Action Workflows", 10 }, { "Action Type", 10 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the MPR summary diffgram dataset.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateManagementPolicyRulesSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("MPRsSummary") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Display Name"); // for sorting purposes
                var column2 = new DataColumn("Display Name Markup");
                var column3 = new DataColumn("Disabled");
                var column4 = new DataColumn("Grant Permission");
                var column5 = new DataColumn("AuthN Workflows");
                var column6 = new DataColumn("AuthZ Workflows");
                var column7 = new DataColumn("Action Workflows");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.Columns.Add(column4);
                table.Columns.Add(column5);
                table.Columns.Add(column6);
                table.Columns.Add(column7);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("MPRActionTypes") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("Display Name");
                var column22 = new DataColumn("Action Type");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.PrimaryKey = new[] { column12, column22 };

                var printTable = Documenter.GetPrintTable();

                // Table 1
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 5 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 6 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                var diffgramTable = Documenter.CreateDiffgramTable(table);
                var diffgramTable2 = Documenter.CreateDiffgramTable(table2);

                this.DiffgramDataSet = new DataSet("MPRsSummary") { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSet.Tables.Add(diffgramTable);
                this.DiffgramDataSet.Tables.Add(diffgramTable2);
                this.DiffgramDataSet.Tables.Add(printTable);

                // set up data relations
                var dataRelation12 = new DataRelation("DataRelation12", new[] { diffgramTable.Columns[0] }, new[] { diffgramTable2.Columns[0] }, false);
                this.DiffgramDataSet.Relations.Add(dataRelation12);

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
        /// Fills the MPRs summary diffgram dataset
        /// </summary>
        protected void FillManagementPolicyRulesSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];
                var diffgramTable2 = this.DiffgramDataSet.Tables[1];

                var objectType = "ManagementPolicyRule";

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
                        var attributeObjectIdChange = this.GetAttributeChange("ObjectID");
                        var displayNameChange = this.GetAttributeChange("DisplayName");
                        var disabledChange = this.GetAttributeChange("Disabled");
                        var grantRightChange = this.GetAttributeChange("GrantRight");
                        var hasAuthNWorkflowsChange = this.TestAttribute("AuthenticationWorkflowDefinition");
                        var hasAuthZWorkflowsChange = this.TestAttribute("AuthorizationWorkflowDefinition");
                        var hasActionWorkflowsChange = this.TestAttribute("ActionWorkflowDefinition");
                        var actionTypeChange = this.GetAttributeChange("ActionType");

                        var objectId = attributeObjectIdChange.NewValue;
                        var displayName = displayNameChange.NewValue;
                        var displayNameMarkup = ServiceCommonDocumenter.GetJumpToBookmarkLocationMarkup(displayName, objectId, objectModificationType);
                        var disabled = disabledChange.NewValue;
                        var disabledOld = disabledChange.OldValue;
                        var grantRight = grantRightChange.NewValue;
                        var grantRightOld = grantRightChange.OldValue;
                        var hasAuthNWorkflows = hasAuthNWorkflowsChange.NewValue;
                        var hasAuthNWorkflowsOld = hasAuthNWorkflowsChange.OldValue;
                        var hasAuthZWorkflows = hasAuthZWorkflowsChange.NewValue;
                        var hasAuthZWorkflowsOld = hasAuthZWorkflowsChange.OldValue;
                        var hasActionWorkflows = hasActionWorkflowsChange.NewValue;
                        var hasActionWorkflowsOld = hasActionWorkflowsChange.OldValue;

                        Documenter.AddRow(diffgramTable, new object[] { displayName, displayNameMarkup, disabled, grantRight, hasAuthNWorkflows, hasAuthZWorkflows, hasActionWorkflows, objectModificationType, displayNameMarkup, disabledOld, grantRightOld, hasAuthNWorkflowsOld, hasAuthZWorkflowsOld, hasActionWorkflowsOld });

                        foreach (var actionTypeValueChange in actionTypeChange.AttributeValues)
                        {
                            Documenter.AddRow(diffgramTable2, new object[] { displayName, actionTypeValueChange.NewValue, actionTypeValueChange.ValueModificationType });
                        }
                    }

                    this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion MPRs Summary

        #region MPR Details

        /// <summary>
        /// Processes the MPR detailed configuration.
        /// </summary>
        protected void ProcessManagementPolicyRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Management Policy Rules";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "ManagementPolicyRule";

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
                        this.ProcessManagementPolicyRule();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current MPR object.
        /// </summary>
        protected void ProcessManagementPolicyRule()
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
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Type", "ManagementPolicyRuleType"),
                    new KeyValuePair<string, string>("Disabled", "Disabled"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                var managementPolicyRuleTypeChange = this.GetAttributeChange("ManagementPolicyRuleType");

                if (managementPolicyRuleTypeChange.NewValue == "Request")
                {
                    // Requestors and Operations
                    this.CreateNestedMultivalueOrderedSettingsDiffgramDataSet();
                    this.FillManagementPolicyRulesRequestorsAndOperationsDiffgramDataSet();
                    this.PrintManagementPolicyRulesRequestorsAndOperationsDiffgramDataSet();

                    // Target Resources
                    this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                    this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                    {
                    new KeyValuePair<string, string>("Target Resource Definition Before Request", "ResourceCurrentSet"),
                    new KeyValuePair<string, string>("Target Resource Definition After Request", "ResourceFinalSet"),
                    });
                    this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Target Resources", 30 }, { string.Empty, 70 } });

                    this.CreateNestedMultivalueOrderedSettingsDiffgramDataSet();
                    this.FillManagementPolicyRulesTargetResourceAttributesDiffgramDataSet();
                    this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { string.Empty, 30 }, { " ", 30 }, { "   ", 40 } });
                }
                else
                {
                    // Transition Definition
                    this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                    this.FillManagementPolicyRulesTransitionDefinitionDiffgramDataSet();
                    this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Transition Definition", 30 }, { string.Empty, 70 } });
                }

                // Policy Workflows
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Authentication", "AuthenticationWorkflowDefinition"),
                    new KeyValuePair<string, string>("Authorization", "AuthorizationWorkflowDefinition"),
                    new KeyValuePair<string, string>("Action", "ActionWorkflowDefinition"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Policy Workflows", 30 }, { string.Empty, 70 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region MPR - Requestors and Operations

        /// <summary>
        /// Fills the MPR Requestors and Operations diffgram dataset
        /// </summary>
        protected void FillManagementPolicyRulesRequestorsAndOperationsDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];
                var diffgramTable2 = this.DiffgramDataSet.Tables[1];
                var diffgramTable3 = this.DiffgramDataSet.Tables[2];
                var configurationIndex = 0;

                var principalSetChange = this.GetAttributeChange("PrincipalSet");
                var principalRelativeToResourceChange = this.GetAttributeChange("PrincipalRelativeToResource");

                Documenter.AddRow(diffgramTable, new object[] { "Requestors", DataRowState.Unchanged });

                Documenter.AddRow(diffgramTable2, new object[] { "Requestors", "Specific Set of Requestors", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable3, new object[] { "Requestors", "Specific Set of Requestors", ++configurationIndex, principalSetChange.NewValueText, principalSetChange.NewValue, principalSetChange.AttributeModificationType, principalSetChange.OldValueText, principalSetChange.OldValue });

                Documenter.AddRow(diffgramTable2, new object[] { "Requestors", "Relative to Resource", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable3, new object[] { "Requestors", "Relative to Resource", ++configurationIndex, principalRelativeToResourceChange.NewValueText, principalRelativeToResourceChange.NewValue, principalRelativeToResourceChange.AttributeModificationType, principalRelativeToResourceChange.OldValueText, principalRelativeToResourceChange.OldValue });

                var actionTypeChange = this.GetAttributeChange("ActionType");
                AttributeValueChange createResourceChange = this.GetManagementPolicyRuleActionTypeChange(actionTypeChange, "Create");
                AttributeValueChange deleteResourceChange = this.GetManagementPolicyRuleActionTypeChange(actionTypeChange, "Delete");
                AttributeValueChange readResourceChange = this.GetManagementPolicyRuleActionTypeChange(actionTypeChange, "Read");
                AttributeValueChange addResourceChange = this.GetManagementPolicyRuleActionTypeChange(actionTypeChange, "Add");
                AttributeValueChange removeResourceChange = this.GetManagementPolicyRuleActionTypeChange(actionTypeChange, "Remove");
                AttributeValueChange modifyResourceChange = this.GetManagementPolicyRuleActionTypeChange(actionTypeChange, "Modify");

                Documenter.AddRow(diffgramTable, new object[] { "Operations", DataRowState.Unchanged });

                Documenter.AddRow(diffgramTable2, new object[] { "Operations", "Create resource", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable3, new object[] { "Operations", "Create resource", ++configurationIndex, createResourceChange.NewValueText, createResourceChange.NewValue, createResourceChange.ValueModificationType, createResourceChange.OldValueText, createResourceChange.OldValue });

                Documenter.AddRow(diffgramTable2, new object[] { "Operations", "Delete resource", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable3, new object[] { "Operations", "Delete resource", ++configurationIndex, deleteResourceChange.NewValueText, deleteResourceChange.NewValue, deleteResourceChange.ValueModificationType, deleteResourceChange.OldValueText, deleteResourceChange.OldValue });

                Documenter.AddRow(diffgramTable2, new object[] { "Operations", "Read resource", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable3, new object[] { "Operations", "Read resource", ++configurationIndex, readResourceChange.NewValueText, readResourceChange.NewValue, readResourceChange.ValueModificationType, readResourceChange.OldValueText, readResourceChange.OldValue });

                Documenter.AddRow(diffgramTable2, new object[] { "Operations", "Add a value to a multivalued attribute", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable3, new object[] { "Operations", "Add a value to a multivalued attribute", ++configurationIndex, addResourceChange.NewValueText, addResourceChange.NewValue, addResourceChange.ValueModificationType, addResourceChange.OldValueText, addResourceChange.OldValue });

                Documenter.AddRow(diffgramTable2, new object[] { "Operations", "Remove a value from a multivalued attribute", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable3, new object[] { "Operations", "Remove a value from a multivalued attribute", ++configurationIndex, removeResourceChange.NewValueText, removeResourceChange.NewValue, removeResourceChange.ValueModificationType, removeResourceChange.OldValueText, removeResourceChange.OldValue });

                Documenter.AddRow(diffgramTable2, new object[] { "Operations", "Modify a single-valued attribute", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable3, new object[] { "Operations", "Modify a single-valued attribute", ++configurationIndex, modifyResourceChange.NewValueText, modifyResourceChange.NewValue, modifyResourceChange.ValueModificationType, modifyResourceChange.OldValueText, modifyResourceChange.OldValue });

                var grantRightChange = this.GetAttributeChange("GrantRight");
                var grantRight = grantRightChange.NewValue == "True" ? "Yes" : "No";
                var grantRightOld = grantRightChange.OldValue == "True" ? "Yes" : "No";

                Documenter.AddRow(diffgramTable, new object[] { "Permissions", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable2, new object[] { "Permissions", "Grants permission", DataRowState.Unchanged });
                Documenter.AddRow(diffgramTable3, new object[] { "Permissions", "Grants permission", ++configurationIndex, grantRight, grantRight, grantRightChange.AttributeModificationType, grantRightOld, grantRightOld });

                this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets if there was a change to MPR for the specified action type.
        /// </summary>
        /// <param name="actionTypeChange">The ActionType change.</param>
        /// <param name="actionType">The action type values to test the change for.</param>
        /// <returns>True, if the specified action type was changed.</returns>
        protected AttributeValueChange GetManagementPolicyRuleActionTypeChange(AttributeChange actionTypeChange, string actionType)
        {
            var operationChange = new AttributeValueChange();
            operationChange.NewValue = "No";
            operationChange.OldValue = "No";
            operationChange.ValueModificationType = actionTypeChange == null || actionTypeChange.AttributeModificationType == DataRowState.Modified ? DataRowState.Unchanged : actionTypeChange.AttributeModificationType;

            if (actionTypeChange != null)
            {
                foreach (var valueChange in actionTypeChange.AttributeValues)
                {
                    if (valueChange.NewValue == actionType)
                    {
                        operationChange.ValueModificationType = valueChange.ValueModificationType;
                        switch (operationChange.ValueModificationType)
                        {
                            case DataRowState.Added:
                                operationChange.NewValue = "Yes";
                                break;
                            case DataRowState.Deleted:
                                operationChange.OldValue = "Yes";
                                break;
                            default:
                                operationChange.NewValue = "Yes";
                                operationChange.OldValue = "Yes";
                                break;
                        }

                        break;
                    }
                }
            }

            return operationChange;
        }

        /// <summary>
        /// Prints the MPR Requestors and Operations diffgram dataset
        /// </summary>
        protected void PrintManagementPolicyRulesRequestorsAndOperationsDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetHeaderTable();
                    headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", "Requestors and Operations" }, { "RowSpan", 1 }, { "ColSpan", 3 } }).Values.Cast<object>().ToArray());

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion MPR - Requestors and Operations

        #region MPR - Target Resources

        /// <summary>
        /// Fills the MPR target resource attributes diffgram dataset
        /// </summary>
        protected void FillManagementPolicyRulesTargetResourceAttributesDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];
                var diffgramTable2 = this.DiffgramDataSet.Tables[1];
                var diffgramTable3 = this.DiffgramDataSet.Tables[2];
                var configurationIndex = 0;

                var actionParameterChange = this.GetAttributeChange("ActionParameter");
                var allAttributesChange = new AttributeValueChange();
                allAttributesChange.NewValue = actionParameterChange.NewValue == "*" ? "Yes" : "No";
                allAttributesChange.OldValue = actionParameterChange.OldValue == "*" ? "Yes" : "No";
                var state = this.CurrentChangeObjectState;
                var objectModificationType = state == "Create" ? DataRowState.Added : state == "Delete" ? DataRowState.Deleted : DataRowState.Modified;
                allAttributesChange.ValueModificationType = objectModificationType != DataRowState.Modified ? objectModificationType : allAttributesChange.NewValue != allAttributesChange.OldValue ? DataRowState.Modified : DataRowState.Unchanged;

                Documenter.AddRow(diffgramTable, new object[] { "Resource Attributes", DataRowState.Unchanged });

                Documenter.AddRow(diffgramTable2, new object[] { "Resource Attributes", "All Attributes", DataRowState.Unchanged });

                Documenter.AddRow(diffgramTable3, new object[] { "Resource Attributes", "All Attributes", ++configurationIndex, allAttributesChange.NewValueText, allAttributesChange.NewValue, allAttributesChange.ValueModificationType, allAttributesChange.OldValueText, allAttributesChange.OldValue });

                Documenter.AddRow(diffgramTable2, new object[] { "Resource Attributes", "Specific Attributes", DataRowState.Unchanged });
                for (var attributeValueIndex = 0; attributeValueIndex < actionParameterChange.AttributeValues.Count; ++attributeValueIndex)
                {
                    Documenter.AddRow(diffgramTable3, new object[] { "Resource Attributes", "Specific Attributes", ++configurationIndex + attributeValueIndex, actionParameterChange.AttributeValues[attributeValueIndex].NewValueText, actionParameterChange.AttributeValues[attributeValueIndex].NewValue, actionParameterChange.AttributeValues[attributeValueIndex].ValueModificationType, actionParameterChange.AttributeValues[attributeValueIndex].OldValueText, actionParameterChange.AttributeValues[attributeValueIndex].OldValue });
                }

                this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion MPR - Target Resources

        #region MPR - Transition Definition

        /// <summary>
        /// Fills the MPR transition definition diffgram dataset
        /// </summary>
        protected void FillManagementPolicyRulesTransitionDefinitionDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var actionTypeChange = this.GetAttributeChange("ActionType");
                var resourceFinalSetChange = this.GetAttributeChange("ResourceFinalSet");
                var resourceCurrentSetChange = this.GetAttributeChange("ResourceCurrentSet");

                this.AddSimpleMultivalueOrderedRows("Transition Set", actionTypeChange.NewValue == "TransitionIn" ? resourceFinalSetChange : resourceCurrentSetChange);
                this.AddSimpleMultivalueOrderedRows("Transition Type", actionTypeChange);

                this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion MPR - Transition Definition

        #endregion MPR Details

        #endregion MPRs

        #region Workflows

        #region Workflows Summary

        /// <summary>
        /// Processes the workflows summary.
        /// </summary>
        protected void ProcessWorkflowsSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Workflows and Activities Summary";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                this.CreateSimpleSummarySettingsDiffgramDataSet(4);
                this.FillWorkflowsSummaryDiffgramDataSet();
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Display Name", 40 }, { "Workflow Type", 20 }, { "Activities", 40 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the workflows summary diffgram dataset
        /// </summary>
        protected void FillWorkflowsSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];
                var diffgramTable2 = this.DiffgramDataSet.Tables[1];

                var objectType = "WorkflowDefinition";

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
                        var attributeObjectIdChange = this.GetAttributeChange("ObjectID");
                        var displayNameChange = this.GetAttributeChange("DisplayName");
                        var requestPhaseChange = this.GetAttributeChange("RequestPhase");

                        var objectId = attributeObjectIdChange.NewValue;
                        var displayName = displayNameChange.NewValue;
                        var displayNameMarkup = ServiceCommonDocumenter.GetJumpToBookmarkLocationMarkup(displayName, objectId, objectModificationType);
                        var requestPhase = requestPhaseChange.NewValue;
                        var requestPhaseOld = requestPhaseChange.OldValue;

                        var xomlChange = this.GetAttributeChange("XOML");
                        var xoml = xomlChange.NewValueText;
                        var xomlOld = xomlChange.OldValueText;

                        var activityList = string.IsNullOrEmpty(xoml) ? string.Empty : string.Join("<br/>", XElement.Parse(xoml).Elements().Select(e => e.Name.LocalName));
                        var activityListOld = string.IsNullOrEmpty(xomlOld) ? string.Empty : string.Join("<br/>", XElement.Parse(xomlOld).Elements().Select(e => e.Name.LocalName));

                        Documenter.AddRow(diffgramTable, new object[] { displayName, displayNameMarkup, requestPhase, activityList, objectModificationType, displayNameMarkup, requestPhaseOld, activityListOld });
                    }

                    this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Workflows Summary

        #region Worflows Details

        /// <summary>
        /// Processes the workflows detailed configuration.
        /// </summary>
        protected void ProcessWorkflows()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Workflows";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "WorkflowDefinition";

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
                        this.ProcessWorkflow();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current workflow object.
        /// </summary>
        protected void ProcessWorkflow()
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
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Workflow Type", "RequestPhase"),
                    new KeyValuePair<string, string>("Run On Policy Update", "RunOnPolicyUpdate"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                // Workflow Activties
                this.CreateWorkflowActivityDetailsDataSets();
                this.FillWorkflowActivityDetails(true);
                this.FillWorkflowActivityDetails(false);

                this.CreateWorkflowActivityAssemblyVersionDiffgramDataSet();
                this.CreateWorkflowActivitySelectedOptionsDiffgramDataSet();
                this.CreateWorkflowActivityMultiValuesDiffgramDataSet();
                this.CreateWorkflowActivityValueExpressionsDiffgramDataSet();
                this.CreateWorkflowActivityUpdateExpressionsDiffgramDataSet();

                foreach (DataRow row in this.DiffgramDataSets[0].Tables[0].Rows)
                {
                    this.WriteBreakTag();

                    var activityIndex = (int)row["ActivityIndex"];
                    var activityType = (string)row["Activity Type"];
                    this.PrintWorkflowAssemblyVersion(activityIndex);
                    this.PrintWorkflowSelectedOptions(activityIndex);
                    switch (activityType)
                    {
                        case "QAAuthenticationGate":
                            this.PrintWorkflowActivityMultiValuesTable(activityIndex, new OrderedDictionary { { "QAGate Questions", 100 } });
                            break;
                        case "VerifyRequest":
                            this.PrintWorkflowActivityValueExpressionsTable(activityIndex, 1, "Contitions", new OrderedDictionary { { "Required Condition for the Request", 50 }, { "Denial Message if Condition is not Satisfied", 50 } });
                            break;
                        case "CreateResource":
                            this.PrintWorkflowActivityValueExpressionsTable(activityIndex, 1, "Query Resources", new OrderedDictionary { { "Query Key", 20 }, { "XPath Filter", 80 } });
                            this.PrintWorkflowActivityValueExpressionsTable(activityIndex, 2, "Attributes", new OrderedDictionary { { "Source Expression", 70 }, { "Target", 30 } });
                            break;
                        case "UpdateResources":
                            this.PrintWorkflowActivityValueExpressionsTable(activityIndex, 1, "Query Resources", new OrderedDictionary { { "Query Key", 20 }, { "XPath Filter", 80 } });
                            this.PrintWorkflowActivityValueExpressionsTable(activityIndex, 2, "Updates", new OrderedDictionary { { "Source Expression", 70 }, { "Target", 25 }, { "Allow Null", 5 } });
                            break;
                        case "GenerateUniqueValue":
                            this.PrintWorkflowActivityValueExpressionsTable(activityIndex, 1, "LDAP Queries", new OrderedDictionary { { "Directory Entry Path", 35 }, { "LDAP Filter", 65 } });
                            this.PrintWorkflowActivityMultiValuesTable(activityIndex, new OrderedDictionary { { "UniqueValue Expression", 100 } });
                            break;
                        case "RunPowerShellScript":
                            this.PrintWorkflowActivityValueExpressionsTable(activityIndex, 1, "Parameters", new OrderedDictionary { { "Parameter", 30 }, { "Value Expression", 70 } });
                            this.PrintWorkflowActivityMultiValuesTable(activityIndex, new OrderedDictionary { { "Arguments", 100 } });
                            break;
                        case "DeleteResources":
                        case "AddDelay":
                        case "SendEmailNotification":
                        case "RequestApproval":
                        case "PWResetActivity":
                        case "ApprovalActivity":
                        case "FilterValidationActivity":
                        case "FunctionActivity":
                        case "GroupValidationActivity":
                        case "EmailNotificationActivity":
                        case "RequestorValidationActivity":
                        case "SynchronizationRuleActivity":
                        case "CreateResourceActivity":
                        case "ReadResourceActivity":
                        case "UpdateResourceActivity":
                        case "DeleteResourceActivity":
                        case "PAMRequestHandlerActivity":
                        case "AddUserToGroupActivity":
                        case "PAMRequestValidationActivity":
                        case "PAMRequestMFASequenceActivity":
                        case "PAMRequestApprovalSequenceActivity":
                        case "PAMRequestDelayedValidationSequenceActivity":
                        case "PAMRequestAvailabilityWindowValidationActivity":
                        case "PasswordCheckGate":
                        case "LockoutGate":
                            break;
                        default:
                            this.PrintWorkflowActivityUnhandledWarning(activityType);
                            break;
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Workflow Activity Details

        /// <summary>
        /// Creates the workflow activity details datasets.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateWorkflowActivityDetailsDataSets()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("Workflow Activities") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("ActivityIndex", typeof(int));
                var column2 = new DataColumn("Activity Type");
                var column3 = new DataColumn("Activity Display Name");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.PrimaryKey = new[] { column1, column2 };

                var table2 = new DataTable("Workflow Activity Assembly Version") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("ActivityIndex", typeof(int));
                var column22 = new DataColumn("Assembly Version");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.PrimaryKey = new[] { column12 };

                var table3 = new DataTable("Workflow Activity Selected Options") { Locale = CultureInfo.InvariantCulture };

                var column13 = new DataColumn("ActivityIndex", typeof(int));
                var column23 = new DataColumn("Selected Option");
                var column33 = new DataColumn("ValueMarkup");
                var column43 = new DataColumn("ValueText");

                table3.Columns.Add(column13);
                table3.Columns.Add(column23);
                table3.Columns.Add(column33);
                table3.Columns.Add(column43);
                table3.PrimaryKey = new[] { column13, column23 };

                var table4 = new DataTable("Configuration Multivalues") { Locale = CultureInfo.InvariantCulture }; // for QA Gate Questions

                var column14 = new DataColumn("ActivityIndex", typeof(int));
                var column24 = new DataColumn("ConfigurationIndex");
                var column34 = new DataColumn("Configuration");

                table4.Columns.Add(column14);
                table4.Columns.Add(column24);
                table4.Columns.Add(column34);
                table4.PrimaryKey = new[] { column14, column24 };

                var table5 = new DataTable("Workflow Activity WAL Value Expressions") { Locale = CultureInfo.InvariantCulture };

                var column15 = new DataColumn("ActivityIndex", typeof(int));
                var column25 = new DataColumn("SectionIndex", typeof(int)); // QueryTable, UpdatesTable, etc...
                var column35 = new DataColumn("Value Expression");
                var column45 = new DataColumn("Target");

                table5.Columns.Add(column15);
                table5.Columns.Add(column25);
                table5.Columns.Add(column35);
                table5.Columns.Add(column45);
                table5.PrimaryKey = new[] { column15, column25, column35, column45 };

                var table6 = new DataTable("Workflow Activity WAL Update Expressions") { Locale = CultureInfo.InvariantCulture };

                var column16 = new DataColumn("ActivityIndex", typeof(int));
                var column26 = new DataColumn("SectionIndex", typeof(int)); // QueryTable, UpdatesTable, etc...
                var column36 = new DataColumn("Value Expression");
                var column46 = new DataColumn("Target");
                var column56 = new DataColumn("Allow Null");

                table6.Columns.Add(column16);
                table6.Columns.Add(column26);
                table6.Columns.Add(column36);
                table6.Columns.Add(column46);
                table6.Columns.Add(column56);
                table6.PrimaryKey = new[] { column16, column26, column36, column46 };

                this.PilotDataSet = new DataSet("Workflow Activities") { Locale = CultureInfo.InvariantCulture };
                this.PilotDataSet.Tables.Add(table);
                this.PilotDataSet.Tables.Add(table2);
                this.PilotDataSet.Tables.Add(table3);
                this.PilotDataSet.Tables.Add(table4);
                this.PilotDataSet.Tables.Add(table5);
                this.PilotDataSet.Tables.Add(table6);

                var dataRelation12 = new DataRelation("DataRelation12", new[] { column1 }, new[] { column12 }, false);
                var dataRelation13 = new DataRelation("DataRelation13", new[] { column1 }, new[] { column13 }, false);
                var dataRelation14 = new DataRelation("DataRelation14", new[] { column1 }, new[] { column14 }, false);
                var dataRelation15 = new DataRelation("DataRelation15", new[] { column1 }, new[] { column15 }, false);
                var dataRelation16 = new DataRelation("DataRelation16", new[] { column1 }, new[] { column16 }, false);

                this.PilotDataSet.Relations.Add(dataRelation12);
                this.PilotDataSet.Relations.Add(dataRelation13);
                this.PilotDataSet.Relations.Add(dataRelation14);
                this.PilotDataSet.Relations.Add(dataRelation15);
                this.PilotDataSet.Relations.Add(dataRelation16);

                this.ProductionDataSet = this.PilotDataSet.Clone();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the workflows activity details
        /// </summary>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillWorkflowActivityDetails(bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry("Pilot Config: '{0}'.", pilotConfig);

            try
            {
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;
                var activityBaseTable = dataSet.Tables[0];
                var activityAssemblyVersionTable = dataSet.Tables[1];
                var activitySelectedOptionsTable = dataSet.Tables[2];
                var activityConfigurationMultivaluesTable = dataSet.Tables[3];
                var activityValueExpressionsTable = dataSet.Tables[4];
                var activityUpdateExpressionsTable = dataSet.Tables[5];

                var xomlChange = this.GetAttributeChange("XOML");
                var xomlXml = ServiceCommonDocumenter.ConvertXomlToSimplifiedXml(pilotConfig ? xomlChange.NewValueText : xomlChange.OldValueText);

                if (xomlXml != null)
                {
                    var activities = xomlXml.Elements();
                    for (var activityIndex = 0; activityIndex < activities.Count(); ++activityIndex)
                    {
                        var activity = activities.ElementAt(activityIndex);
                        var activityType = activity.Name.LocalName;

                        if (activityType == "AuthenticationGateActivity")
                        {
                            activity = activity.XPathSelectElement("AuthenticationGateActivity.AuthenticationGate/child::node()");
                            activityType = activity.Name.LocalName;
                        }

                        var assemblyVersion = activity.Attribute("AssemblyVersion");
                        var activityDisplayName = (string)activity.Attribute("ActivityDisplayName");
                        activityDisplayName = string.IsNullOrEmpty(activityDisplayName) ? activityType : activityDisplayName + " (" + activityType + ")";

                        Documenter.AddRow(activityBaseTable, new object[] { activityIndex, activityType, activityDisplayName });
                        Documenter.AddRow(activityAssemblyVersionTable, new object[] { activityIndex, assemblyVersion });

                        foreach (var attribute in activity.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration))
                        {
                            if (attribute.Name == "AssemblyVersion")
                            {
                                continue;
                            }

                            var selectedOption = attribute.Name.LocalName;
                            var optionValue = attribute.Value;
                            var optionValueMarkup = attribute.Value;

                            if (ServiceCommonDocumenter.IsGuid(optionValue))
                            {
                                var valueModificationType = this.Environment == ConfigEnvironment.PilotOnly ? DataRowState.Added : this.Environment == ConfigEnvironment.ProductionOnly ? DataRowState.Deleted : DataRowState.Unchanged; // Or should it be modified?
                                var resolvedValue = TryResolveReferences(optionValue, pilotConfig, valueModificationType);
                                optionValueMarkup = resolvedValue.Item1;
                                optionValue = resolvedValue.Item2;
                            }

                            if (optionValue != "{x:Null}" && optionValue != "00000000-0000-0000-0000-000000000000" && selectedOption != "Name" && selectedOption != "ActivityDisplayName")
                            {
                                Documenter.AddRow(activitySelectedOptionsTable, new object[] { activityIndex, selectedOption, optionValueMarkup, optionValue });
                            }
                        }

                        switch (activityType)
                        {
                            case "QAAuthenticationGate":
                                this.FillWorkflowActivityExpressionsList(activityConfigurationMultivaluesTable, activity.XPathSelectElement("QAAuthenticationGate.Questions/Array"), activityIndex);
                                break;
                            case "VerifyRequest":
                                this.FillWorkflowActivityValueExpressions(activityValueExpressionsTable, activity.XPathSelectElement("VerifyRequest.ConditionsTable/Hashtable"), activityIndex, 1);
                                break;
                            case "CreateResource":
                                this.FillWorkflowActivityValueExpressions(activityValueExpressionsTable, activity.XPathSelectElement("CreateResource.QueriesTable/Hashtable"), activityIndex, 1);
                                this.FillWorkflowActivityValueExpressions(activityValueExpressionsTable, activity.XPathSelectElement("CreateResource.AttributesTable/Hashtable"), activityIndex, 2);
                                break;
                            case "UpdateResources":
                                this.FillWorkflowActivityValueExpressions(activityValueExpressionsTable, activity.XPathSelectElement("UpdateResources.QueriesTable/Hashtable"), activityIndex, 1);
                                this.FillWorkflowActivityValueExpressions(activityUpdateExpressionsTable, activity.XPathSelectElement("UpdateResources.UpdatesTable/Hashtable"), activityIndex, 2);
                                break;
                            case "GenerateUniqueValue":
                                this.FillWorkflowActivityValueExpressions(activityValueExpressionsTable, activity.XPathSelectElement("GenerateUniqueValue.LdapQueriesTable/Hashtable"), activityIndex, 1);
                                this.FillWorkflowActivityExpressionsList(activityConfigurationMultivaluesTable, activity.XPathSelectElement("GenerateUniqueValue.ValueExpressions/ArrayList"), activityIndex);
                                break;
                            case "RunPowerShellScript":
                                this.FillWorkflowActivityValueExpressions(activityValueExpressionsTable, activity.XPathSelectElement("RunPowerShellScript.ParametersTable/Hashtable"), activityIndex, 1);
                                this.FillWorkflowActivityExpressionsList(activityConfigurationMultivaluesTable, activity.XPathSelectElement("RunPowerShellScript.Arguments/ArrayList"), activityIndex);
                                break;
                        }
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Pilot Config: '{0}'.", pilotConfig);
            }
        }

        /// <summary>
        /// Fills the value expressions of a WAL workflow activity
        /// </summary>
        /// <param name="activityValueExpressionsTable">The value expression table to fill.</param>
        /// <param name="hashtable">The hashtable element of the activity XOML.</param>
        /// <param name="activityIndex">The index of the activity in the workflow.</param>
        /// <param name="sectionIndex">The index of the section to which the table belongs.</param>
        protected void FillWorkflowActivityValueExpressions(DataTable activityValueExpressionsTable, XContainer hashtable, int activityIndex, int sectionIndex)
        {
            if (activityValueExpressionsTable != null && hashtable != null)
            {
                var hashtableKeyCount = hashtable.Nodes().Count() / 3;

                for (var expressionIndex = 0; expressionIndex < hashtableKeyCount; ++expressionIndex)
                {
                    var valueExpression = ((IEnumerable)hashtable.XPathEvaluate("//Key[String = '" + expressionIndex + ":0']/parent::node()/text()")).Cast<XText>().FirstOrDefault();
                    var target = ((IEnumerable)hashtable.XPathEvaluate("//Key[String = '" + expressionIndex + ":1']/parent::node()/text()")).Cast<XText>().FirstOrDefault();
                    var allowNull = ((IEnumerable)hashtable.XPathEvaluate("//Key[String = '" + expressionIndex + ":2']/parent::node()/text()")).Cast<XText>().FirstOrDefault();

                    if (activityValueExpressionsTable.Columns.Count == 4)
                    {
                        Documenter.AddRow(activityValueExpressionsTable, new object[] { activityIndex, sectionIndex, valueExpression != null ? valueExpression.Value : null, target != null ? target.Value : null });
                    }
                    else
                    {
                        Documenter.AddRow(activityValueExpressionsTable, new object[] { activityIndex, sectionIndex, valueExpression != null ? valueExpression.Value : null, target != null ? target.Value : null, allowNull != null ? allowNull.Value : null });
                    }
                }
            }
        }

        /// <summary>
        /// Fills the expressions list a workflow activity
        /// </summary>
        /// <param name="activityConfigurationMultivaluesTable">The multi-values list table to fill.</param>
        /// <param name="list">The Array or ArrayList element of the activity XOML.</param>
        /// <param name="activityIndex">The index of the current activity in the workflow.</param>
        protected void FillWorkflowActivityExpressionsList(DataTable activityConfigurationMultivaluesTable, XContainer list, int activityIndex)
        {
            Logger.Instance.WriteMethodEntry();
            try
            {
                if (list != null)
                {
                    var elements = list.Elements();

                    for (var configIndex = 0; configIndex < elements.Count(); ++configIndex)
                    {
                        Documenter.AddRow(activityConfigurationMultivaluesTable, new object[] { activityIndex, configIndex, (string)elements.ElementAt(configIndex) });
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Creates the workflow activity assembly version diffgram dataset.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateWorkflowActivityAssemblyVersionDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = Documenter.GetDiffgram(this.PilotDataSet.Tables[0], this.ProductionDataSet.Tables[0], null);
                var diffgramTable2 = Documenter.GetDiffgram(this.PilotDataSet.Tables[1], this.ProductionDataSet.Tables[1], null);

                var printTable = Documenter.GetPrintTable();

                // Table 2
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                this.DiffgramDataSet = new DataSet("Workflow Activities") { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSet.Tables.Add(diffgramTable);
                this.DiffgramDataSet.Tables.Add(diffgramTable2);
                this.DiffgramDataSet.Tables.Add(printTable);

                // set up data relations
                var dataRelation12 = new DataRelation("DataRelation12", new[] { diffgramTable.Columns[0] }, new[] { diffgramTable2.Columns[0] }, false);
                this.DiffgramDataSet.Relations.Add(dataRelation12);

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
        /// Creates the workflow activity selected options diffgram dataset.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateWorkflowActivitySelectedOptionsDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var printTable = Documenter.GetPrintTable();

                // Table 2
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", true } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", true }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                var columnsIgnored = printTable.Select("ChangeIgnored = true AND TableIndex = 1").Select(row => (int)row["ColumnIndex"]).ToArray();

                var diffgramTable = Documenter.GetDiffgram(this.PilotDataSet.Tables[0], this.ProductionDataSet.Tables[0], null);
                var diffgramTable2 = Documenter.GetDiffgram(this.PilotDataSet.Tables[2], this.ProductionDataSet.Tables[2], columnsIgnored);

                this.DiffgramDataSet = new DataSet("Workflow Activities") { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSet.Tables.Add(diffgramTable);
                this.DiffgramDataSet.Tables.Add(diffgramTable2);
                this.DiffgramDataSet.Tables.Add(printTable);

                // set up data relations
                var dataRelation12 = new DataRelation("DataRelation12", new[] { diffgramTable.Columns[0] }, new[] { diffgramTable2.Columns[0] }, false);
                this.DiffgramDataSet.Relations.Add(dataRelation12);

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
        /// Creates the workflow activity multi-values diffgram dataset.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateWorkflowActivityMultiValuesDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = Documenter.GetDiffgram(this.PilotDataSet.Tables[0], this.ProductionDataSet.Tables[0], null);
                var diffgramTable2 = Documenter.GetDiffgram(this.PilotDataSet.Tables[3], this.ProductionDataSet.Tables[3], null);

                var printTable = Documenter.GetPrintTable();

                // Table 2
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                this.DiffgramDataSet = new DataSet("Workflow Activity Multi-values") { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSet.Tables.Add(diffgramTable);
                this.DiffgramDataSet.Tables.Add(diffgramTable2);
                this.DiffgramDataSet.Tables.Add(printTable);

                // set up data relations
                var dataRelation12 = new DataRelation("DataRelation12", new[] { diffgramTable.Columns[0] }, new[] { diffgramTable2.Columns[0] }, false);
                this.DiffgramDataSet.Relations.Add(dataRelation12);

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
        /// Creates the workflow activity value expressions diffgram dataset.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateWorkflowActivityValueExpressionsDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = Documenter.GetDiffgram(this.PilotDataSet.Tables[0], this.ProductionDataSet.Tables[0], null);
                var diffgramTable2 = Documenter.GetDiffgram(this.PilotDataSet.Tables[4], this.ProductionDataSet.Tables[4], null);

                var printTable = Documenter.GetPrintTable();

                // Table 2
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                this.DiffgramDataSet = new DataSet("Workflow Activity Multi-values") { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSet.Tables.Add(diffgramTable);
                this.DiffgramDataSet.Tables.Add(diffgramTable2);
                this.DiffgramDataSet.Tables.Add(printTable);

                // set up data relations
                var dataRelation12 = new DataRelation("DataRelation12", new[] { diffgramTable.Columns[0] }, new[] { diffgramTable2.Columns[0] }, false);
                this.DiffgramDataSet.Relations.Add(dataRelation12);

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
        /// Creates the workflow activity update expressions diffgram dataset.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateWorkflowActivityUpdateExpressionsDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = Documenter.GetDiffgram(this.PilotDataSet.Tables[0], this.ProductionDataSet.Tables[0], null);
                var diffgramTable2 = Documenter.GetDiffgram(this.PilotDataSet.Tables[5], this.ProductionDataSet.Tables[5], null);

                var printTable = Documenter.GetPrintTable();

                // Table 2
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                this.DiffgramDataSet = new DataSet("Workflow Activity Multi-values") { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSet.Tables.Add(diffgramTable);
                this.DiffgramDataSet.Tables.Add(diffgramTable2);
                this.DiffgramDataSet.Tables.Add(printTable);

                // set up data relations
                var dataRelation12 = new DataRelation("DataRelation12", new[] { diffgramTable.Columns[0] }, new[] { diffgramTable2.Columns[0] }, false);
                this.DiffgramDataSet.Relations.Add(dataRelation12);

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
        /// Prints the workflow activity assembly version information
        /// </summary>
        /// <param name="activityIndex">The activity index</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void PrintWorkflowAssemblyVersion(int activityIndex)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = this.DiffgramDataSets[0];
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var currentActivityRow = this.DiffgramDataSet.Tables[0].Rows[activityIndex];
                    var activityName = (string)currentActivityRow[2];
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable((activityIndex + 1) + ". " + activityName, new OrderedDictionary { { "Activity Assembly Version", 100 } });

                    #region table

                    this.ReportWriter.WriteBeginTag("table");
                    this.ReportWriter.WriteAttribute("class", HtmlTableSize.Standard.ToString() + " " + this.GetCssVisibilityClass());
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        this.WriteTableHeader(headerTable);
                    }

                    #region rows

                    int currentTableIndex = 1;
                    int currentCellIndex = 0;
                    this.WriteRows(this.DiffgramDataSet.Tables[currentTableIndex].Select("ActivityIndex = " + activityIndex), currentTableIndex, ref currentCellIndex);

                    #endregion rows

                    this.ReportWriter.WriteEndTag("table");
                    this.ReportWriter.WriteLine();
                    this.ReportWriter.Flush();

                    #endregion table
                }
            }
            finally
            {
                ////this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the workflow activity selected options
        /// </summary>
        /// <param name="activityIndex">The activity index</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void PrintWorkflowSelectedOptions(int activityIndex)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = this.DiffgramDataSets[1];

                int currentTableIndex = 1;
                var rows = this.DiffgramDataSet.Tables[currentTableIndex].Select("ActivityIndex = " + activityIndex);
                if (rows.Count() != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Selected Option", 30 }, { "Value", 70 } });

                    #region table

                    this.ReportWriter.WriteBeginTag("table");
                    this.ReportWriter.WriteAttribute("class", HtmlTableSize.Standard.ToString() + " " + this.GetCssVisibilityClass());
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        this.WriteTableHeader(headerTable);
                    }

                    #region rows

                    int currentCellIndex = 0;
                    this.WriteRows(rows, currentTableIndex, ref currentCellIndex);

                    #endregion rows

                    this.ReportWriter.WriteEndTag("table");
                    this.ReportWriter.WriteLine();
                    this.ReportWriter.Flush();

                    #endregion table
                }
            }
            finally
            {
                ////this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the workflow activity multi-values table
        /// </summary>
        /// <param name="activityIndex">The activity index.</param>
        /// <param name="tableHeader">The table header.</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void PrintWorkflowActivityMultiValuesTable(int activityIndex, OrderedDictionary tableHeader)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.DiffgramDataSet = this.DiffgramDataSets[2];

                int currentTableIndex = 1;
                var rows = this.DiffgramDataSet.Tables[currentTableIndex].Select("ActivityIndex = " + activityIndex);

                if (rows.Count() != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(tableHeader);

                    #region table

                    this.ReportWriter.WriteBeginTag("table");
                    this.ReportWriter.WriteAttribute("class", HtmlTableSize.Standard.ToString() + " " + this.GetCssVisibilityClass());
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        this.WriteTableHeader(headerTable);
                    }

                    #region rows

                    int currentCellIndex = 0;
                    this.WriteRows(rows, currentTableIndex, ref currentCellIndex);

                    #endregion rows

                    this.ReportWriter.WriteEndTag("table");
                    this.ReportWriter.WriteLine();
                    this.ReportWriter.Flush();

                    #endregion table
                }
            }
            finally
            {
                ////this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the workflow activity value expressions table
        /// </summary>
        /// <param name="activityIndex">The activity index.</param>
        /// <param name="sectionIndex">The section index.</param>
        /// <param name="tableHeader">The table header.</param>
        /// <param name="tableSubHeader">The table sub header.</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void PrintWorkflowActivityValueExpressionsTable(int activityIndex, int sectionIndex, string tableHeader, OrderedDictionary tableSubHeader)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (tableSubHeader == null)
                {
                    return;
                }

                this.DiffgramDataSet = tableSubHeader.Count == 3 ? this.DiffgramDataSets[4] : this.DiffgramDataSets[3];

                int currentTableIndex = 1;
                var rows = this.DiffgramDataSet.Tables[currentTableIndex].Select("ActivityIndex = " + activityIndex + "AND SectionIndex = " + sectionIndex);
                if (rows.Count() != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable(tableHeader, tableSubHeader);

                    #region table

                    this.ReportWriter.WriteBeginTag("table");
                    this.ReportWriter.WriteAttribute("class", HtmlTableSize.Standard.ToString() + " " + this.GetCssVisibilityClass());
                    this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                    {
                        this.WriteTableHeader(headerTable);
                    }

                    #region rows

                    int currentCellIndex = 0;
                    this.WriteRows(rows, currentTableIndex, ref currentCellIndex);

                    #endregion rows

                    this.ReportWriter.WriteEndTag("table");
                    this.ReportWriter.WriteLine();
                    this.ReportWriter.Flush();

                    #endregion table
                }
            }
            finally
            {
                ////this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the workflow activity unhandled warning
        /// </summary>
        /// <param name="activityType">The activity type.</param>
        [SuppressMessage("StyleCop.CSharp.ReadabilityRules", "SA1123:DoNotPlaceRegionsWithinElements", Justification = "Reviewed.")]
        protected void PrintWorkflowActivityUnhandledWarning(string activityType)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var warning = string.Format(CultureInfo.InvariantCulture, "!!WARNING!! The activity '{0}' is not handled by the documenter. The documentation of this activity may be incomplete.", activityType);

                Logger.Instance.WriteWarning(warning);

                this.ReportWriter.WriteBeginTag("table");
                this.ReportWriter.WriteAttribute("class", HtmlTableSize.Standard.ToString());
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                this.ReportWriter.WriteFullBeginTag("tr");
                this.ReportWriter.WriteBeginTag("td");
                this.ReportWriter.WriteAttribute("class", "Highlight");
                this.ReportWriter.Write(HtmlTextWriter.TagRightChar);
                this.ReportWriter.Write(warning);
                this.ReportWriter.WriteEndTag("td");
                this.ReportWriter.WriteEndTag("tr");
                this.ReportWriter.WriteEndTag("table");
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Workflow Activity Details

        #endregion Worflows Details

        #endregion Workflows

        #region Synchronization Rules

        #region Synchronization Rules Summary

        /// <summary>
        /// Processes the Synchronization Rules summary.
        /// </summary>
        protected void ProcessSynchronizationRulesSummary()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Synchronization Rules Summary";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                this.CreateSimpleSummarySettingsDiffgramDataSet(5);
                this.FillSynchronizationRulesSummaryDiffgramDataSet();
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Display Name", 40 }, { "Precedence", 10 }, { "Connected System", 20 }, { "Description", 30 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the synchronization rules summary diffgram dataset
        /// </summary>
        protected void FillSynchronizationRulesSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];

                var objectType = "SynchronizationRule";

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
                        var attributeObjectIdChange = this.GetAttributeChange("ObjectID");
                        var displayNameChange = this.GetAttributeChange("DisplayName");
                        var precedenceChange = this.GetAttributeChange("Precedence");
                        var managementAgentIdChange = this.GetAttributeChange("ManagementAgentID");
                        var descriptionChange = this.GetAttributeChange("Description");

                        var objectId = attributeObjectIdChange.NewValue;
                        var displayName = displayNameChange.NewValue;
                        var displayNameMarkup = ServiceCommonDocumenter.GetJumpToBookmarkLocationMarkup(displayName, objectId, objectModificationType);
                        var precedence = precedenceChange.NewValue;
                        var precedenceOld = precedenceChange.OldValue;
                        var managementAgentId = managementAgentIdChange.NewValue;
                        var managementAgentIdOld = managementAgentIdChange.OldValue;
                        var description = descriptionChange.NewValue;
                        var descriptionOld = descriptionChange.OldValue;

                        Documenter.AddRow(diffgramTable, new object[] { displayName, displayNameMarkup, precedence, managementAgentId, description, objectModificationType, displayNameMarkup, precedenceOld, managementAgentIdOld, descriptionOld });
                    }

                    this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Synchronization Rules Summary

        #region Synchronization Rules Details

        /// <summary>
        /// Processes the synchronization rules detailed configuration.
        /// </summary>
        protected void ProcessSynchronizationRules()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Synchronization Rules";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "SynchronizationRule";

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
                        this.ProcessSynchronizationRule();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current synchronization rule object.
        /// </summary>
        protected void ProcessSynchronizationRule()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.PrintSimpleSectionHeader(4);

                // General Info
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSynchronizationRuleGeneralInfoDiffgramDataSet();
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                // Scope
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Metaverse Resource Type", "ILMObjectType"),
                    new KeyValuePair<string, string>("External System", "ManagementAgentID"),
                    new KeyValuePair<string, string>("External System Resource Type", "ConnectedObjectType"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Scope", 30 }, { string.Empty, 70 } });

                // Outbound System Scoping Filter
                var outboundScopingFiltersChange = this.GetAttributeChange("msidmOutboundScopingFilters");
                this.ProcessSynchronizationRuleSystemScopingFilter(outboundScopingFiltersChange);

                // Inbound System Scoping Filter
                var inboundScopingFiltersChange = this.GetAttributeChange("ConnectedSystemScope");
                this.ProcessSynchronizationRuleSystemScopingFilter(inboundScopingFiltersChange);

                // Relationship
                var headerTable = Documenter.GetSimpleSettingsHeaderTable(new OrderedDictionary { { "Relationship", 100 } });
                this.WriteTable(new DataTable(), headerTable);

                // Relationship Criteria
                var relationshipCriteriaChange = this.GetAttributeChange("RelationshipCriteria");
                this.ProcessSynchronizationRuleRelationshipCriteria(relationshipCriteriaChange);

                // Provisioning Configuration
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Create Resource in FIM", "CreateILMObject"),
                    new KeyValuePair<string, string>("Create Resource in External System", "CreateConnectedSystemObject"),
                    new KeyValuePair<string, string>("Enable Deprovisioning", "DisconnectConnectedSystemObject"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Provisioning Configuration", 70 }, { string.Empty, 30 } });

                // Workflow Parameters
                var synchronizationRuleParametersChange = this.GetAttributeChange("SynchronizationRuleParameters");
                this.ProcessSynchronizationRuleWorkflowParameters(synchronizationRuleParametersChange);

                // Outbound Attribute Flows
                this.ProcessSynchronizationRuleOutboundTransformations();

                // Inbound Attribute Flows
                this.ProcessSynchronizationRuleInboundTransformations();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the Synchronization Rule general info diffgram dataset
        /// </summary>
        protected void FillSynchronizationRuleGeneralInfoDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.AddSimpleMultivalueOrderedRows("Resource Type", this.GetAttributeChange("ObjectType"));
                this.AddSimpleMultivalueOrderedRows("Display Name", this.GetAttributeChange("DisplayName"));
                this.AddSimpleMultivalueOrderedRows("Description", this.GetAttributeChange("Description"));
                this.AddSimpleMultivalueOrderedRows("Dependency", this.GetAttributeChange("Dependency"));

                var flowTypeChange = this.GetAttributeChange("FlowType");
                foreach (var attributeValue in flowTypeChange.AttributeValues)
                {
                    attributeValue.NewValue = (attributeValue.NewValue == "0") ? "Inbound" : (attributeValue.NewValue == "1") ? "Outbound" : (attributeValue.NewValue == "2") ? "Inbound and Outbound" : attributeValue.NewValue;
                    attributeValue.OldValue = (attributeValue.OldValue == "0") ? "Inbound" : (attributeValue.OldValue == "1") ? "Outbound" : (attributeValue.OldValue == "2") ? "Inbound and Outbound" : attributeValue.OldValue;
                }

                this.AddSimpleMultivalueOrderedRows("Data Flow Direction", flowTypeChange);

                var outboundRuleIsFilterBasedChange = this.GetAttributeChange("msidmOutboundIsFilterBased");
                foreach (var attributeValueChange in outboundRuleIsFilterBasedChange.AttributeValues)
                {
                    if (!string.IsNullOrEmpty(attributeValueChange.NewValue) && flowTypeChange.NewValue != "Inbound")
                    {
                        attributeValueChange.NewValue = attributeValueChange.NewValue.Equals("False", StringComparison.OrdinalIgnoreCase) ? "Policy-based Outbound Synchronization" : "Filter-based Outbound Synchronization";
                    }
                    else
                    {
                        attributeValueChange.NewValue = string.Empty;
                    }

                    if (!string.IsNullOrEmpty(attributeValueChange.OldValue) && flowTypeChange.OldValue != "Inbound")
                    {
                        attributeValueChange.OldValue = attributeValueChange.OldValue.Equals("False", StringComparison.OrdinalIgnoreCase) ? "Policy-based Outbound Synchronization" : "Filter-based Outbound Synchronization";
                    }
                    else
                    {
                        attributeValueChange.OldValue = string.Empty;
                    }
                }

                this.AddSimpleMultivalueOrderedRows("Apply Rule", outboundRuleIsFilterBasedChange);

                this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the Synchronization Rule Inbound / Outbound System Scoping Filters
        /// </summary>
        /// <param name="systemScopingFiltersChange">The inbound / outbound system scoping filter attribute change.</param>
        protected void ProcessSynchronizationRuleSystemScopingFilter(AttributeChange systemScopingFiltersChange)
        {
            if (systemScopingFiltersChange == null)
            {
                return;
            }

            this.CreateSimpleOrderedSettingsDataSets(4, 3, false); // 1 = Display Order Control, 2 = Attribute, 3 = Operator, 4 = Value
            this.FillSynchronizationRuleSystemScopingFiltersDataSet(systemScopingFiltersChange.NewValue, true);
            this.FillSynchronizationRuleSystemScopingFiltersDataSet(systemScopingFiltersChange.OldValue, false);
            this.CreateSimpleOrderedSettingsDiffgram();
            this.PrintSynchronizationRuleSystemScopingFilters(true);
        }

        /// <summary>
        /// Fills the Synchronization Rule Inbound / Outbound System Scoping Filters dataset
        /// </summary>
        /// <param name="systemScopingFiltersXml">The inbound / outbound system scoping filter xml.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillSynchronizationRuleSystemScopingFiltersDataSet(string systemScopingFiltersXml, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;
                var table = dataSet.Tables[0];

                if (!string.IsNullOrEmpty(systemScopingFiltersXml))
                {
                    var outboundScopingFilters = XElement.Parse(systemScopingFiltersXml).XPathSelectElements("//scope");
                    for (var i = 0; i < outboundScopingFilters.Count(); ++i)
                    {
                        var outboundScopingFilter = outboundScopingFilters.ElementAt(i);
                        var csAttribute = (string)outboundScopingFilter.Element("csAttribute");
                        var csOperator = (string)outboundScopingFilter.Element("csOperator");
                        var csValue = (string)outboundScopingFilter.Element("csValue");
                        Documenter.AddRow(table, new object[] { i, csAttribute, csOperator, csValue });
                    }

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the Synchronization Rule Inbound / Outbound System Scoping Filters
        /// </summary>
        /// <param name="outboundSystemScopingFilters">if set to <c>true</c>, denotes outbound system scoping filter. Otherwise, the .</param>
        protected void PrintSynchronizationRuleSystemScopingFilters(bool outboundSystemScopingFilters)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetHeaderTable();
                    headerTable.Rows.Add((new OrderedDictionary { { "RowIndex", 0 }, { "ColumnIndex", 0 }, { "ColumnName", outboundSystemScopingFilters ? "Outbound System Scoping Filter" : "Inbound System Scoping Filter" }, { "RowSpan", 1 }, { "ColSpan", 3 } }).Values.Cast<object>().ToArray());

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the Synchronization Rule Relationship Criteria
        /// </summary>
        /// <param name="relationshipCriteriaChange">The Synchronization Rule Relationship Criteria attribute change.</param>
        protected void ProcessSynchronizationRuleRelationshipCriteria(AttributeChange relationshipCriteriaChange)
        {
            if (relationshipCriteriaChange == null)
            {
                return;
            }

            this.CreateSimpleOrderedSettingsDataSets(4, 3, false); // 1 = Display Order Control, 2 = Attribute, 3 = Operator, 4 = Value
            this.FillSynchronizationRuleRelationshipCriteriaDataSet(relationshipCriteriaChange.NewValue, true);
            this.FillSynchronizationRuleRelationshipCriteriaDataSet(relationshipCriteriaChange.OldValue, false);
            this.CreateSimpleOrderedSettingsDiffgram();
            this.PrintSynchronizationRuleRelationshipCriteria();
        }

        /// <summary>
        /// Fills the Synchronization Rule Relationship Criteria dataset
        /// </summary>
        /// <param name="relationshipCriteriaXml">The sync rule relationship criteria xml.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillSynchronizationRuleRelationshipCriteriaDataSet(string relationshipCriteriaXml, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;
                var table = dataSet.Tables[0];

                if (!string.IsNullOrEmpty(relationshipCriteriaXml))
                {
                    var outboundScopingFilters = XElement.Parse(relationshipCriteriaXml).XPathSelectElements("//condition");
                    for (var i = 0; i < outboundScopingFilters.Count(); ++i)
                    {
                        var outboundScopingFilter = outboundScopingFilters.ElementAt(i);
                        var ilmAttribute = (string)outboundScopingFilter.Element("ilmAttribute");
                        var csAttribute = (string)outboundScopingFilter.Element("csAttribute");
                        Documenter.AddRow(table, new object[] { i, ilmAttribute, "=", csAttribute });
                    }

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the Synchronization Rule Relationship Criteria
        /// </summary>
        protected void PrintSynchronizationRuleRelationshipCriteria()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable("Relationship Criteria", new OrderedDictionary { { "Metaverse Object Attribute", 45 }, { "=", 10 }, { "Connected System Object Attribute", 45 } });
                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the Synchronization Rule Workflow Parameters
        /// </summary>
        /// <param name="synchronizationRuleParametersChange">The  Synchronization Rule Workflow Parameters attribute change.</param>
        protected void ProcessSynchronizationRuleWorkflowParameters(AttributeChange synchronizationRuleParametersChange)
        {
            if (synchronizationRuleParametersChange == null)
            {
                return;
            }

            this.CreateSimpleSettingsDataSets(2); // 1 = Parameter Name, 2 = Parameter Type
            this.FillSynchronizationRuleWorkflowParametersDataSet(synchronizationRuleParametersChange.NewValue, true);
            this.FillSynchronizationRuleWorkflowParametersDataSet(synchronizationRuleParametersChange.OldValue, false);
            this.CreateSimpleOrderedSettingsDiffgram();
            this.PrintSynchronizationRuleWorkflowParameters();
        }

        /// <summary>
        /// Fills the Synchronization Rule Workflow Parameters dataset
        /// </summary>
        /// <param name="synchronizationRuleParametersXml">The sync rule workflow parameters xml.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillSynchronizationRuleWorkflowParametersDataSet(string synchronizationRuleParametersXml, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;
                var table = dataSet.Tables[0];

                if (!string.IsNullOrEmpty(synchronizationRuleParametersXml))
                {
                    var outboundScopingFilters = XElement.Parse(synchronizationRuleParametersXml).XPathSelectElements("//sync-parameter");
                    foreach (var outboundScopingFilter in outboundScopingFilters)
                    {
                        var parameterName = (string)outboundScopingFilter.Element("name");
                        var parameterType = (string)outboundScopingFilter.Element("type");
                        Documenter.AddRow(table, new object[] { parameterName, parameterType });
                    }

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the Synchronization Rule Workflow Parameters
        /// </summary>
        protected void PrintSynchronizationRuleWorkflowParameters()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable("Workflow Parameters", new OrderedDictionary { { "Name", 70 }, { "Data Type", 30 } });
                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Parses the Synchronization Rule Function Expression
        /// </summary>
        /// <param name="function">The function to be parsed.</param>
        /// <returns>The parsed function expression string.</returns>
        protected string ParseSynchronizationRuleFunctionExpression(XElement function)
        {
            if (function == null)
            {
                return string.Empty;
            }

            var customExpression = ((string)function.Attribute("isCustomExpression") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? true : false;
            var functionName = (string)function.Attribute("id") ?? string.Empty;

            var functionExpression = string.Empty;
            if (functionName == "+")
            {
                foreach (var arg in function.Elements("arg"))
                {
                    if (arg.Elements("fn").Count() > 0)
                    {
                        functionExpression += this.ParseSynchronizationRuleFunctionExpression(arg.Element("fn"));
                    }
                    else
                    {
                        functionExpression += (string)arg;
                    }

                    functionExpression += "+";
                }

                functionExpression = functionExpression.TrimEnd('+');
            }
            else
            {
                functionExpression = functionName + "(";
                foreach (var arg in function.Elements("arg"))
                {
                    if (arg.Elements("fn").Count() > 0)
                    {
                        functionExpression += this.ParseSynchronizationRuleFunctionExpression(arg.Element("fn"));
                    }
                    else
                    {
                        functionExpression += (string)arg;
                    }

                    functionExpression += ",";
                }

                functionExpression = functionExpression.TrimEnd(',') + ")";
            }

            if (customExpression)
            {
                functionExpression = "CustomExpression(" + functionExpression + ")";
            }

            return functionExpression;
        }

        /// <summary>
        /// Processes the Synchronization Rule Outbound Transformations
        /// </summary>
        protected void ProcessSynchronizationRuleOutboundTransformations()
        {
            this.CreateSimpleOrderedSettingsDataSets(8, 2, true); // 1 = Display Order Control (Destination Attribute), 2 = Initial Flow, 3 = Existence Test, 4 = Allow Null, 5 = Reference Attribute Precedence, 6 = FIM Value, 7 = Direction, 8 = Destination Attribute

            var synchronizationRuleInitialFlowChange = this.GetAttributeChange("InitialFlow");
            var synchronizationRulePersistentFlowChange = this.GetAttributeChange("PersistentFlow");
            var synchronizationRuleExistenceTestChange = this.GetAttributeChange("ExistenceTest");

            this.FillSynchronizationRuleOutboundTransformationsDataSet(synchronizationRuleInitialFlowChange, true);
            this.FillSynchronizationRuleOutboundTransformationsDataSet(synchronizationRuleInitialFlowChange, false);
            this.FillSynchronizationRuleOutboundTransformationsDataSet(synchronizationRulePersistentFlowChange, true);
            this.FillSynchronizationRuleOutboundTransformationsDataSet(synchronizationRulePersistentFlowChange, false);
            this.FillSynchronizationRuleOutboundTransformationsDataSet(synchronizationRuleExistenceTestChange, true);
            this.FillSynchronizationRuleOutboundTransformationsDataSet(synchronizationRuleExistenceTestChange, false);
            this.CreateSimpleOrderedSettingsDiffgram();
            this.PrintSynchronizationRuleOutboundTransformations();
        }

        /// <summary>
        /// Fills the Synchronization Rule Outbound Transformations dataset
        /// </summary>
        /// <param name="transformationsChange">The sync rule transformations xml.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillSynchronizationRuleOutboundTransformationsDataSet(AttributeChange transformationsChange, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;
                var table = dataSet.Tables[0];

                if (transformationsChange != null && transformationsChange.AttributeValues != null && transformationsChange.AttributeValues.Count > 0)
                {
                    foreach (var valueChange in transformationsChange.AttributeValues)
                    {
                        var exportFlowXml = pilotConfig ? valueChange.NewValue : valueChange.OldValue;
                        if (!string.IsNullOrEmpty(exportFlowXml))
                        {
                            var exportFlow = XElement.Parse(exportFlowXml);
                            var destination = (string)exportFlow.Element("dest");
                            var initialFlow = transformationsChange.AttributeName.Equals("InitialFlow", StringComparison.OrdinalIgnoreCase) ? true : false;
                            var existenceTest = transformationsChange.AttributeName.Equals("ExistenceTest", StringComparison.OrdinalIgnoreCase) ? true : false;
                            var allowNull = ((string)exportFlow.Attribute("allows-null") ?? string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase) ? true : false;
                            var referenceAttributePrecedence = string.Empty;
                            foreach (var csValue in exportFlow.XPathSelectElements("scope/csValue"))
                            {
                                referenceAttributePrecedence += (string)csValue + ";";
                            }

                            var directFlow = exportFlow.Elements("fn").Count() == 0;
                            var source = string.Empty;
                            if (directFlow)
                            {
                                source = exportFlow.XPathSelectElement("src/attr") != null ? (string)exportFlow.XPathSelectElement("src/attr") : (string)exportFlow.Element("src");
                            }
                            else
                            {
                                source = this.ParseSynchronizationRuleFunctionExpression(exportFlow.Element("fn"));
                            }

                            Documenter.AddRow(table, new object[] { destination, initialFlow, existenceTest, allowNull, referenceAttributePrecedence, source, "&#8594;", destination });
                        }
                    }

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the Synchronization Rule Outbound Transformations
        /// </summary>
        protected void PrintSynchronizationRuleOutboundTransformations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable("Outbound Attribute Flow", new OrderedDictionary { { "Initial Flow Only", 10 }, { "Use as Existence Test", 10 }, { "Allow Null", 5 }, { "Reference Attribute Precedence", 15 }, { "FIM Value", 35 }, { "Direction", 10 }, { "Destination Attribute", 15 } });
                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the Synchronization Rule Inbound Transformations
        /// </summary>
        protected void ProcessSynchronizationRuleInboundTransformations()
        {
            this.CreateSimpleOrderedSettingsDataSets(4, 1, true); // 1 = Display Order Control (FIM Attribute), 2 = External System Attributes / Values, 3 = Direction, 4 = FIM Attribute

            var synchronizationRulePersistentFlowChange = this.GetAttributeChange("PersistentFlow");

            this.FillSynchronizationRuleInboundTransformationsDataSet(synchronizationRulePersistentFlowChange, true);
            this.FillSynchronizationRuleInboundTransformationsDataSet(synchronizationRulePersistentFlowChange, false);
            this.CreateSimpleOrderedSettingsDiffgram();
            this.PrintSynchronizationRuleInboundTransformations();
        }

        /// <summary>
        /// Fills the Synchronization Rule Inbound Transformations dataset
        /// </summary>
        /// <param name="transformationsChange">The sync rule transformations xml.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, the pilot configuration is loaded. Otherwise, the production configuration is loaded.</param>
        protected void FillSynchronizationRuleInboundTransformationsDataSet(AttributeChange transformationsChange, bool pilotConfig)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var dataSet = pilotConfig ? this.PilotDataSet : this.ProductionDataSet;
                var table = dataSet.Tables[0];

                if (transformationsChange != null && transformationsChange.AttributeValues != null && transformationsChange.AttributeValues.Count > 0)
                {
                    foreach (var valueChange in transformationsChange.AttributeValues)
                    {
                        var importFlowXml = pilotConfig ? valueChange.NewValue : valueChange.OldValue;
                        if (!string.IsNullOrEmpty(importFlowXml))
                        {
                            var importFlow = XElement.Parse(importFlowXml);
                            var destination = (string)importFlow.Element("dest");
                            var directFlow = importFlow.Elements("fn").Count() == 0;
                            var source = string.Empty;
                            if (directFlow)
                            {
                                source = importFlow.XPathSelectElement("src/attr") != null ? (string)importFlow.XPathSelectElement("src/attr") : (string)importFlow.Element("src");
                            }
                            else
                            {
                                source = this.ParseSynchronizationRuleFunctionExpression(importFlow.Element("fn"));
                            }

                            Documenter.AddRow(table, new object[] { destination, source, "&#8594;", destination });
                        }
                    }

                    table.AcceptChanges();
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the Synchronization Rule Inbound Transformations
        /// </summary>
        protected void PrintSynchronizationRuleInboundTransformations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable("Inbound Attribute Flow", new OrderedDictionary { { "External System Attributes/Values", 65 }, { "Direction", 10 }, { "FIM Attribute", 25 } });
                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Synchronization Rules Details

        #endregion Synchronization Rules

        #region Email Templates

        /// <summary>
        /// Processes the EmailTemplate objects.
        /// </summary>
        protected void ProcessEmailTemplates()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Email Templates";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "EmailTemplate";

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
                        this.ProcessEmailTemplate();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current EmailTemplate  object.
        /// </summary>
        protected void ProcessEmailTemplate()
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
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Template Type", "EmailTemplateType"),
                    new KeyValuePair<string, string>("Subject", "EmailSubject"),
                });

                var emailBodyChange = this.GetAttributeChange("EmailBody");
                if (!string.IsNullOrEmpty(emailBodyChange.OldValue))
                {
                    emailBodyChange.AttributeValues[0].OldValue = WebUtility.HtmlEncode(emailBodyChange.OldValue);
                }

                if (!string.IsNullOrEmpty(emailBodyChange.NewValue))
                {
                    emailBodyChange.AttributeValues[0].NewValue = WebUtility.HtmlEncode(emailBodyChange.NewValue);
                }

                this.AddSimpleMultivalueOrderedRows("Body", emailBodyChange);

                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Email Templates

        #region Resource Control Display Configurations

        /// <summary>
        /// Processes the ObjectVisualizationConfiguration objects.
        /// </summary>
        protected void ProcessObjectVisualizationConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Resource Control Display Configurations";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                this.CreateSimpleSummarySettingsDiffgramDataSet(8);
                this.FillObjectVisualizationConfigurationsSummaryDiffgramDataSet();
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Display Name", 30 }, { "Target Resource Type", 20 }, { "Applies to Create", 5 }, { "Applies to Edit", 5 }, { "Applies to View", 5 }, { "Configuration XML Change", 10 }, { "Description", 25 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Fills the ObjectVisualizationConfiguration diffgram dataset
        /// </summary>
        protected void FillObjectVisualizationConfigurationsSummaryDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];

                var objectType = "ObjectVisualizationConfiguration";

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
                        var displayNameChange = this.GetAttributeChange("DisplayName");
                        var targetObjectTypeChange = this.GetAttributeChange("TargetObjectType");
                        var appliesToCreateChange = this.GetAttributeChange("AppliesToCreate");
                        var appliesToEditChange = this.GetAttributeChange("AppliesToEdit");
                        var appliesToViewChange = this.GetAttributeChange("AppliesToView");
                        var configurationDataChange = this.GetAttributeChange("ConfigurationData");
                        var descriptionChange = this.GetAttributeChange("Description");

                        var displayName = displayNameChange.NewValue;
                        var targetObjectType = targetObjectTypeChange.NewValue;
                        var targetObjectTypeOld = targetObjectTypeChange.OldValue;
                        var appliesToCreate = appliesToCreateChange.NewValue;
                        var appliesToCreateOld = appliesToCreateChange.OldValue;
                        var appliesToEdit = appliesToEditChange.NewValue;
                        var appliesToEditOld = appliesToEditChange.OldValue;
                        var appliesToView = appliesToViewChange.NewValue;
                        var appliesToViewOld = appliesToViewChange.OldValue;
                        var configurationData = configurationDataChange.AttributeModificationType == DataRowState.Modified ? "New RCDC XML" : "RCDC XML";
                        var configurationDataOld = configurationDataChange.AttributeModificationType == DataRowState.Modified ? "Old RCDC XML" : "RCDC XML";
                        var description = descriptionChange.NewValue;
                        var descriptionOld = descriptionChange.OldValue;

                        Documenter.AddRow(diffgramTable, new object[] { displayName, displayName, targetObjectType, appliesToCreate, appliesToEdit, appliesToView, configurationData, description, objectModificationType, displayName, targetObjectTypeOld, appliesToCreateOld, appliesToEditOld, appliesToViewOld, configurationDataOld, description });
                    }

                    this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Resource Control Display Configurations

        #region Portal UI Configuration

        /// <summary>
        /// Processes the PortalUIConfiguration objects.
        /// </summary>
        protected void ProcessPortalUIConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Portal UI Configuration";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "PortalUIConfiguration";

                var changeObjects = this.GetChangedObjects(objectType);

                if (changeObjects.Count() == 0)
                {
                    this.WriteContentParagraph(string.Format(CultureInfo.CurrentUICulture, DocumenterResources.NoCustomizationsDetected, ServiceCommonDocumenter.GetObjectTypeDisplayName(objectType)));
                }
                else
                {
                    // Only one PortalUIConfiguration object is expected at the most.
                    foreach (var changeObject in changeObjects)
                    {
                        this.CurrentChangeObject = changeObject;
                        this.ProcessPortalUIConfiguration();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current PortalUIConfiguration object.
        /// </summary>
        protected void ProcessPortalUIConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                ////this.PrintSimpleSectionHeader(4); // Only one PortalUIConfiguration object is expected at the most.

                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Branding Center Text", "BrandingCenterText"),
                    new KeyValuePair<string, string>("Branding Left Image", "BrandingLeftImage"),
                    new KeyValuePair<string, string>("Branding Right Image", "BrandingRightImage"),
                    new KeyValuePair<string, string>("Global Chache Duration", "UICacheTime"),
                    new KeyValuePair<string, string>("ListView Cache Time Out", "ListViewCacheTimeOut"),
                    new KeyValuePair<string, string>("ListView Items per Page", "ListViewPageSize"),
                    new KeyValuePair<string, string>("istView Pages to Cache", "ListViewPagesToCache"),
                    new KeyValuePair<string, string>("Navigation Bar Resource Count Cache Duration", "UICountCacheTime"),
                    new KeyValuePair<string, string>("Per User Cache Duration", "UIUserCacheTime"),
                    new KeyValuePair<string, string>("Time Zone", "TimeZone"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Setting", 50 }, { "Configuration", 50 } });
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Portal UI Configuration

        #region HomepageConfiguration Configurations

        /// <summary>
        /// Processes the HomepageConfiguration objects.
        /// </summary>
        protected void ProcessHomepageConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Home Page Resources";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "HomepageConfiguration";

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
                        this.ProcessHomepageConfiguration();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current HomepageConfiguration  object.
        /// </summary>
        protected void ProcessHomepageConfiguration()
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
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Image Url", "ImageUrl"),
                    new KeyValuePair<string, string>("Usage Keyword", "UsageKeyword"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                // UI Position
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Region", "Region"),
                    new KeyValuePair<string, string>("Parent Order", "ParentOrder"),
                    new KeyValuePair<string, string>("Order", "Order"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "UI Position", 30 }, { string.Empty, 70 } });

                // Behavior
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Navigation Url", "NavigationUrl"),
                    new KeyValuePair<string, string>("Resource Count", "ResourceCount"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Behavior", 30 }, { string.Empty, 70 } });

                // Localization Info
                this.ProcessLocalizationConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion HomepageConfiguration Configurations

        #region NavigationBarConfiguration Configurations

        /// <summary>
        /// Processes the NavigationBarConfiguration objects.
        /// </summary>
        protected void ProcessNavigationBarConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Navigation Bar Resources";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "NavigationBarConfiguration";

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
                        this.ProcessNavigationBarConfiguration();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current NavigationBarConfiguration  object.
        /// </summary>
        protected void ProcessNavigationBarConfiguration()
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
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Usage Keyword", "UsageKeyword"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                // UI Position
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Parent Order", "ParentOrder"),
                    new KeyValuePair<string, string>("Order", "Order"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "UI Position", 30 }, { string.Empty, 70 } });

                // Behavior
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Navigation Url", "NavigationUrl"),
                    new KeyValuePair<string, string>("Resource Count", "ResourceCount"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Behavior", 30 }, { string.Empty, 70 } });

                // Localization Info
                this.ProcessLocalizationConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion NavigationBarConfiguration Configurations

        #region SearchScopeConfiguration Configurations

        /// <summary>
        /// Processes the SearchScopeConfiguration objects.
        /// </summary>
        protected void ProcessSearchScopeConfigurations()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = "Search Scopes";

                Logger.Instance.WriteInfo("Processing " + sectionTitle + ".");

                this.WriteSectionHeader(sectionTitle, 3, ConfigEnvironment.PilotAndProduction);

                var objectType = "SearchScopeConfiguration";

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
                        this.ProcessSearchScopeConfiguration();
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Processes the current SearchScopeConfiguration  object.
        /// </summary>
        protected void ProcessSearchScopeConfiguration()
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
                    new KeyValuePair<string, string>("Display Name", "DisplayName"),
                    new KeyValuePair<string, string>("Description", "Description"),
                    new KeyValuePair<string, string>("Usage Keyword", "UsageKeyword"),
                    new KeyValuePair<string, string>("Order", "Order"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "General", 30 }, { string.Empty, 70 } });

                // Search Definition
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Attribute Searched", "SearchScopeContext"),
                    new KeyValuePair<string, string>("Search Scope Filter", "SearchScope"),
                    new KeyValuePair<string, string>("Search Scope Advance Filter", "msidmSearchScopeAdvancedFilter"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Search Definition", 30 }, { string.Empty, 70 } });

                // Results
                this.CreateSimpleMultivalueOrderedSettingsDiffgramDataSet();
                this.FillSimpleMultivalueOrderedSettingsDiffgramDataSet(new KeyValuePair<string, string>[]
                {
                    new KeyValuePair<string, string>("Resource type", "SearchScopeResultObjectType"),
                    new KeyValuePair<string, string>("Attribute", "SearchScopeColumn"),
                    new KeyValuePair<string, string>("Redirection URL", "SearchScopeTargetURL"),
                });
                this.PrintSimpleSettingsSectionTable(new OrderedDictionary { { "Results", 30 }, { string.Empty, 70 } });

                // Localization Info
                this.ProcessLocalizationConfiguration();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion SearchScopeConfiguration Configurations
    }
}
