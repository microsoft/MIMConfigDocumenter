//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ServiceCommonDocumenter.cs" company="Microsoft">
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
    using System.Collections;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Web.UI;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// The ServiceCommonDocumenter is the base class for FIMService configuration documenters.
    /// </summary>
    public abstract class ServiceCommonDocumenter : Documenter
    {
        /// <summary>
        /// Pilot ExportObject Root XPath
        /// </summary>
        public const string PilotExportObjectRootXPath = "ServiceConfig/Pilot/Results/ExportObject";

        /// <summary>
        /// Production ExportObject Root XPath
        /// </summary>
        public const string ProductionExportObjectRootXPath = "ServiceConfig/Production/Results/ExportObject";

        /// <summary>
        /// Changes ImportObject Root XPath
        /// </summary>
        public const string ChangesImportObjectRootXPath = "ServiceConfig/Changes/Results/ImportObject";

        /// <summary>
        /// Invariant Locale
        /// </summary>
        public const string InvariantLocale = "Invariant";

        /// <summary>
        /// The current ChangeObject
        /// </summary>
        private XElement currentChangeObject = null;

        /// <summary>
        /// The current ChangeObject state
        /// </summary>
        private string currentChangeObjectState = null;

        /// <summary>
        /// The current ChangeObject object type.
        /// </summary>
        private string currentChangeObjectType = null;

        /// <summary>
        /// Gets or sets the current ChangeObject.
        /// </summary>
        /// <value>
        /// The current ChangeObject.
        /// </value>
        protected XElement CurrentChangeObject
        {
            get
            {
                return this.currentChangeObject;
            }

            set
            {
                this.currentChangeObject = value;
                if (this.currentChangeObject != null)
                {
                    var state = (string)this.currentChangeObject.Element("State");
                    var objectType = (string)this.currentChangeObject.Element("ObjectType");
                    this.currentChangeObjectState = state;
                    this.currentChangeObjectType = objectType;
                    this.Environment = state == "Delete" ? ConfigEnvironment.ProductionOnly : state == "Create" ? ConfigEnvironment.PilotOnly : ConfigEnvironment.PilotAndProduction;
                }
                else
                {
                    this.Environment = ConfigEnvironment.PilotAndProduction;
                }
            }
        }

        /// <summary>
        /// Gets the current change object state
        /// </summary>
        protected string CurrentChangeObjectState
        {
            get
            {
                return this.currentChangeObjectState;
            }
        }

        /// <summary>
        /// Gets the current change object type
        /// </summary>
        protected string CurrentChangeObjectType
        {
            get
            {
                return this.currentChangeObjectType;
            }
        }

        /// <summary>
        /// Gets or sets the changes XML.
        /// </summary>
        /// <value>
        /// The changes XML.
        /// </value>
        protected XElement ChangesXml { get; set; }

        /// <summary>
        /// Gets the display name of the given ObjectType.
        /// </summary>
        /// <param name="objectType">The ObjectType of the object.</param>
        /// <returns>The display name of the given ObjectType.</returns>
        public static string GetObjectTypeDisplayName(string objectType)
        {
            string displayName = "Undefined ObjectType";

            switch (objectType)
            {
                case "ObjectTypeDescription":
                    displayName = "Resource Type";
                    break;
                case "AttributeTypeDescription":
                    displayName = "Attribute";
                    break;
                case "BindingDescription":
                    displayName = "Binding";
                    break;
                case "ActivityInformationConfiguration":
                    displayName = "Activity Information";
                    break;
                case "ForestConfiguration":
                    displayName = "Forest";
                    break;
                case "DomainConfiguration":
                    displayName = "Domain";
                    break;
                case "FilterScope":
                    displayName = "Filter Permission";
                    break;
                case "Set":
                    displayName = "Set";
                    break;
                case "ManagementPolicyRule":
                    displayName = "Management Policy Rule";
                    break;
                case "WorkflowDefinition":
                    displayName = "Workflow";
                    break;
                case "SynchronizationRule":
                    displayName = "Synchronization Rule";
                    break;
                case "EmailTemplate":
                    displayName = "Email Template";
                    break;
                case "ObjectVisualizationConfiguration":
                    displayName = "Resource Control Display";
                    break;
                case "PortalUIConfiguration":
                    displayName = "Portal UI";
                    break;
                case "HomepageConfiguration":
                    displayName = "Homepage Resource";
                    break;
                case "NavigationBarConfiguration":
                    displayName = "Navigation Bar";
                    break;
                case "SearchScopeConfiguration":
                    displayName = "Search Scope";
                    break;
            }

            return displayName;
        }

        /// <summary>
        /// Gets the xpath for locating the specified attribute for the specified object in the specified config.
        /// </summary>
        /// <param name="objectIdentifier">The ObjectID of the object</param>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, indicates that this is a pilot configuration. Otherwise, this is a production configuration.</param>
        /// <returns>The xpath for locating the specified attribute of the specified object in the specified config.</returns>
        protected static string GetAttributeXPath(string objectIdentifier, string attributeName, bool pilotConfig)
        {
            return ServiceCommonDocumenter.GetAttributeXPath(objectIdentifier, attributeName, pilotConfig, ServiceCommonDocumenter.InvariantLocale);
        }

        /// <summary>
        /// Gets the xpath for locating the specified attribute for the specified object in the specified config.
        /// </summary>
        /// <param name="objectIdentifier">The ObjectID of the object</param>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, indicates that this is a pilot configuration. Otherwise, this is a production configuration.</param>
        /// <param name="locale">The locale for the localization configuration.</param>
        /// <returns>The xpath for locating the specified attribute of the specified object in the specified config.</returns>
        protected static string GetAttributeXPath(string objectIdentifier, string attributeName, bool pilotConfig, string locale)
        {
            return ServiceCommonDocumenter.GetObjectXPath(objectIdentifier, pilotConfig) + "/" + ServiceCommonDocumenter.GetRelativeAttributeXPath(attributeName, locale);
        }

        /// <summary>
        /// Gets the xpath for locating the specified attribute of the current change object.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="attributeOperation">The attribute operation.</param>
        /// <returns>The xpath for locating the specified attribute of the current change object.</returns>
        protected static string GetChangeAttributeXPath(string attributeName, DataRowState attributeOperation)
        {
            return ServiceCommonDocumenter.GetChangeAttributeXPath(attributeName, attributeOperation, ServiceCommonDocumenter.InvariantLocale);
        }

        /// <summary>
        /// Gets the xpath for locating the specified attribute of the current change object.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="attributeOperation">The attribute operation.</param>
        /// <param name="locale">The locale for the localization config</param>
        /// <returns>The xpath for locating the specified attribute of the current change object.</returns>
        protected static string GetChangeAttributeXPath(string attributeName, DataRowState attributeOperation, string locale)
        {
            var operation = attributeOperation == DataRowState.Modified ? "Replace" : attributeOperation == DataRowState.Added ? "Add" : attributeOperation == DataRowState.Deleted ? "Delete" : "Unknown";
            return "Changes/ImportChange[AttributeName = '" + attributeName + "' and Operation = '" + operation + (locale != ServiceCommonDocumenter.InvariantLocale ? "' and Locale = '" + locale : string.Empty) + "']";
        }

        /// <summary>
        /// Gets the xpath for the specified object in the specified config.
        /// </summary>
        /// <param name="objectIdentifier">The ObjectID of the object</param>
        /// <param name="pilotConfig">if set to <c>true</c>, indicates that this is a pilot configuration. Otherwise, this is a production configuration.</param>
        /// <returns>The xpath for the specified object in the specified config</returns>
        protected static string GetObjectXPath(string objectIdentifier, bool pilotConfig)
        {
            string xpath = "/ResourceManagementObject[ObjectIdentifier = '" + objectIdentifier + "']";

            if (pilotConfig)
            {
                xpath = ServiceCommonDocumenter.PilotExportObjectRootXPath + xpath;
            }
            else
            {
                xpath = ServiceCommonDocumenter.ProductionExportObjectRootXPath + xpath;
            }

            return xpath;
        }

        /// <summary>
        /// Gets the relative xpath for locating the specified attribute.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <returns>The relative xpath of the specified attribute.</returns>
        protected static string GetRelativeAttributeXPath(string attributeName)
        {
            return ServiceCommonDocumenter.GetRelativeAttributeXPath(attributeName, ServiceCommonDocumenter.InvariantLocale);
        }

        /// <summary>
        /// Gets the relative xpath for locating the specified attribute.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="locale">The locale for the localization configuration.</param>
        /// <returns>The relative xpath of the specified attribute.</returns>
        protected static string GetRelativeAttributeXPath(string attributeName, string locale)
        {
            string xpath = "ResourceManagementAttributes/ResourceManagementAttribute[AttributeName ='" + attributeName + "']";

            if (locale != ServiceCommonDocumenter.InvariantLocale)
            {
                xpath = "/LocalizedResourceManagementAttributes/LocalizedResourceManagementAttribute[Culture = '" + locale + "']/" + xpath;
            }

            return xpath;
        }

        /// <summary>
        /// Gets the markup for jumping to the bookmark location.
        /// </summary>
        /// <param name="displayText">The display text of the bookmark.</param>
        /// <param name="sectionGuid">The section guid.</param>
        /// <param name="rowState">The section modification type.</param>
        /// <returns>The markup for jumping to the bookmark location</returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed.")]
        protected static string GetJumpToBookmarkLocationMarkup(string displayText, string sectionGuid, DataRowState rowState)
        {
            string markup = displayText;
            using (var markupWriter = new XhtmlTextWriter(new StringWriter(CultureInfo.InvariantCulture)))
            {
                Documenter.WriteJumpToBookmarkLocation(markupWriter, displayText, displayText, sectionGuid, rowState.ToString());
                markup = markupWriter.InnerWriter.ToString();
            }

            return markup;
        }

        /// <summary>
        /// Extracts a GUID value from a string.
        /// </summary>
        /// <param name="input">A string value to extract a GUID value from.</param>
        /// <returns>Return a GUID value that is contained in the input string.</returns>
        protected static List<string> ExtractGuid(string input)
        {
            List<string> guids = new List<string>();

            string pattern = @"\{?[a-fA-F\d]{8}-([a-fA-F\d]{4}-){3}[a-fA-F\d]{12}\}?";

            foreach (Match match in Regex.Matches(input, pattern))
            {
                var capture = match.Value;
                if (IsGuid(capture))
                {
                    guids.Add(capture);
                }
            }

            return guids;
        }

        /// <summary>
        /// Checks whether a string is a representation of a GUID.
        /// </summary>
        /// <param name="guid">A string value to test.</param>
        /// <returns>Returns true if the string is a representation of a GUID.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "This is an implementation of Guid.TryParse().")]
        protected static bool IsGuid(string guid)
        {
            Guid result;
            return Guid.TryParse(guid, out result) && result != Guid.Empty;
        }

        /// <summary>
        /// Converts an XOML string to simple XML
        /// </summary>
        /// <param name="workflowXoml">The Workflow XOML string</param>
        /// <returns>The XElement object for the simplified XOML string</returns>
        protected static XElement ConvertXomlToSimplifiedXml(string workflowXoml)
        {
            XElement simplifiedXoml = null;
            if (!string.IsNullOrEmpty(workflowXoml))
            {
                simplifiedXoml = RemoveAllNamespaces(XElement.Parse(workflowXoml));
            }

            return simplifiedXoml;
        }

        /// <summary>
        /// Removes all namespaces from the input xml element
        /// If this is WF XOML, then also add AssemblyVersion attribute to the activity element.
        /// </summary>
        /// <param name="inputElement">The input XElement with namespaces.</param>
        /// <returns>The XElement without any namespaces.</returns>
        protected static XElement RemoveAllNamespaces(XElement inputElement)
        {
            if (inputElement == null)
            {
                return null;
            }

            var stripped = new XElement(inputElement.Name.LocalName);

            foreach (var attribute in inputElement.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration))
            {
                stripped.Add(new XAttribute(attribute.Name.LocalName, attribute.Value));
            }

            var parent = inputElement.Parent;
            if (parent != null)
            {
                var parentLocalName = parent.Name.LocalName;
                if ((parentLocalName == "AuthenticationGateActivity.AuthenticationGate" || parentLocalName == "SequentialWorkflow") && !string.IsNullOrEmpty(inputElement.Name.NamespaceName))
                {
                    var assemblyVersion = inputElement.Name.NamespaceName.Split(new string[] { ";Assembly=" }, StringSplitOptions.RemoveEmptyEntries)[1];
                    stripped.Add(new XAttribute("AssemblyVersion", assemblyVersion));
                }
            }

            if (!inputElement.HasElements)
            {
                stripped.Value = inputElement.Value;
            }
            else
            {
                foreach (var childNode in inputElement.Nodes())
                {
                    if (childNode.NodeType == XmlNodeType.Text)
                    {
                        stripped.Add((XText)childNode);
                    }
                    else if (childNode.NodeType == XmlNodeType.Element)
                    {
                        stripped.Add(RemoveAllNamespaces((XElement)childNode));
                    }
                }
            }

            return stripped;
        }

        /// <summary>
        /// Returns the list of changed objects sorted by their display names.
        /// </summary>
        /// <param name="objectType">The object type of the resources.</param>
        /// <returns>The list of changed objects sorted by their display names.</returns>
        protected IEnumerable<XElement> GetChangedObjects(string objectType)
        {
            var changeObjects = this.ChangesXml.XPathSelectElements(ServiceCommonDocumenter.ChangesImportObjectRootXPath + "[ObjectType = '" + objectType + "' and State != 'Resolve']");

            changeObjects = from changeObject in changeObjects
                            let displayName = this.GetChangeObjectDisplayName(changeObject)
                            orderby displayName
                            select changeObject;

            return changeObjects;
        }

        /// <summary>
        /// Get the display name of the specified change object for sorting purposes.
        /// </summary>
        /// <param name="changeObject">The change object.</param>
        /// <returns>The display name of the specified change object.</returns>
        protected string GetChangeObjectDisplayName(XContainer changeObject)
        {
            if (changeObject == null)
            {
                return string.Empty;
            }

            var sourceObjectIdentifier = (string)changeObject.Element("SourceObjectIdentifier");
            var targetObjectIdentifier = (string)changeObject.Element("TargetObjectIdentifier");

            var displayName = string.Empty;
            if (!string.IsNullOrEmpty(sourceObjectIdentifier))
            {
                var pilotAttribute = this.PilotXml.XPathSelectElement(ServiceCommonDocumenter.GetAttributeXPath(sourceObjectIdentifier, "DisplayName", true, ServiceCommonDocumenter.InvariantLocale));
                if (pilotAttribute != null)
                {
                    displayName = (string)pilotAttribute.Element("Value");
                }
            }
            else
            {
                var productionAttribute = this.ProductionXml.XPathSelectElement(ServiceCommonDocumenter.GetAttributeXPath(targetObjectIdentifier, "DisplayName", false, ServiceCommonDocumenter.InvariantLocale));
                if (productionAttribute != null)
                {
                    displayName = (string)productionAttribute.Element("Value");
                }
            }

            return displayName;
        }

        /// <summary>
        /// Gets the xpath for locating the specified attribute of the current change object in the specified config.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, indicates that this is a pilot configuration. Otherwise, this is a production configuration.</param>
        /// <returns>The xpath for locating the specified attribute of the current change object in the specified config.</returns>
        protected string GetAttributeXPath(string attributeName, bool pilotConfig)
        {
            return this.GetAttributeXPath(attributeName, pilotConfig, ServiceCommonDocumenter.InvariantLocale);
        }

        /// <summary>
        /// Gets the xpath for locating the specified attribute of the current change object in the specified config.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, indicates that this is a pilot configuration. Otherwise, this is a production configuration.</param>
        /// <param name="locale">The locale for the localization configuration.</param>
        /// <returns>The xpath for locating the specified attribute of the current change object in the specified config.</returns>
        protected string GetAttributeXPath(string attributeName, bool pilotConfig, string locale)
        {
            var objectIdentifier = pilotConfig ? (string)this.CurrentChangeObject.Element("SourceObjectIdentifier") : (string)this.CurrentChangeObject.Element("TargetObjectIdentifier");

            return ServiceCommonDocumenter.GetAttributeXPath(objectIdentifier, attributeName, pilotConfig, locale);
        }

        /// <summary>
        /// Tests if the current change object had the specified value change for the specified attribute.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="attributeValue">The value of the attribute.</param>
        /// <param name="attributeOperation">The attribute operation.</param>
        /// <returns>True if the current change object had the specified value change for the specified attribute.</returns>
        protected bool TestAttributeValueChange(string attributeName, string attributeValue, DataRowState attributeOperation)
        {
            return this.TestAttributeValueChange(attributeName, attributeValue, attributeOperation, ServiceCommonDocumenter.InvariantLocale);
        }

        /// <summary>
        /// Tests if the current change object had the specified value change for the specified attribute.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="attributeValue">The value of the attribute.</param>
        /// <param name="attributeOperation">The attribute operation.</param>
        /// <param name="locale">The locale for the localization config</param>
        /// <returns>True if the current change object had the specified value change for the specified attribute.</returns>
        protected bool TestAttributeValueChange(string attributeName, string attributeValue, DataRowState attributeOperation, string locale)
        {
            var operation = attributeOperation == DataRowState.Modified ? "Replace" : attributeOperation == DataRowState.Added ? "Add" : attributeOperation == DataRowState.Deleted ? "Delete" : "Unknown";
            return this.CurrentChangeObject.XPathSelectElement("Changes/ImportChange[AttributeName = '" + attributeName + "' and AttributeValue = '" + attributeValue + "' and Operation = '" + operation + (locale != ServiceCommonDocumenter.InvariantLocale ? "' and Locale = '" + locale : string.Empty + "']")) != null;
        }

        /// <summary>
        /// Test if the attribute exists
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <returns>The attribute value change with NewValue True if the attribute exists in Pilot and OldValue True of the it exists in Production.</returns>
        protected AttributeValueChange TestAttribute(string attributeName)
        {
            var existsInPilot = this.PilotXml.XPathSelectElement(this.GetAttributeXPath(attributeName, true)) != null;
            var existsInProduction = this.ProductionXml.XPathSelectElement(this.GetAttributeXPath(attributeName, true)) != null;

            var valueChange = new AttributeValueChange();

            valueChange.ValueModificationType = existsInPilot != existsInProduction ? DataRowState.Modified : DataRowState.Unchanged;
            valueChange.NewValue = existsInPilot.ToString();
            valueChange.OldValue = existsInProduction.ToString();

            return valueChange;
        }

        /// <summary>
        /// Gets the changes to the configuration of the given attribute.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <returns>The changes to the configuration of the given attribute.</returns>
        protected AttributeChange GetAttributeChange(string attributeName)
        {
            Logger.Instance.WriteMethodEntry("Attribute Name : '{0}'", attributeName);

            try
            {
                var attributeChange = new AttributeChange(attributeName);

                var pilotAttribute = this.PilotXml.XPathSelectElement(this.GetAttributeXPath(attributeName, true));
                var productionAttribute = this.ProductionXml.XPathSelectElement(this.GetAttributeXPath(attributeName, false));

                if (pilotAttribute == null && productionAttribute == null)
                {
                    return attributeChange;
                }

                attributeChange.IsMultivalue = pilotAttribute != null ? (string)pilotAttribute.Element("IsMultiValue") == "true" : productionAttribute != null ? (string)productionAttribute.Element("IsMultiValue") == "true" : false;
                attributeChange.HasReference = pilotAttribute != null ? (string)pilotAttribute.Element("HasReference") == "true" : productionAttribute != null ? (string)productionAttribute.Element("HasReference") == "true" : false;

                if (attributeChange.IsMultivalue)
                {
                    switch (this.CurrentChangeObjectState)
                    {
                        case "Create":
                            {
                                foreach (var value in pilotAttribute.XPathSelectElements("Values/child::node()"))
                                {
                                    var attributeValue = new AttributeValueChange();
                                    attributeChange.AttributeModificationType = DataRowState.Added;
                                    attributeValue.ValueModificationType = attributeChange.AttributeModificationType;
                                    attributeValue.NewValue = (string)value;
                                    if (attributeChange.HasReference)
                                    {
                                        this.TryDereferenceValue(attributeValue);
                                    }

                                    attributeChange.AttributeValues.Add(attributeValue);
                                }
                            }

                            break;
                        case "Delete":
                            {
                                foreach (var value in productionAttribute.XPathSelectElements("Values/child::node()"))
                                {
                                    var attributeValue = new AttributeValueChange();
                                    attributeChange.AttributeModificationType = DataRowState.Deleted;
                                    attributeValue.ValueModificationType = attributeChange.AttributeModificationType;
                                    attributeValue.OldValue = (string)value;
                                    attributeValue.NewValue = attributeValue.OldValue;
                                    if (attributeChange.HasReference)
                                    {
                                        this.TryDereferenceValue(attributeValue);
                                    }
                                }
                            }

                            break;
                        case "Put":
                            {
                                attributeChange.AttributeModificationType = DataRowState.Modified;

                                if (pilotAttribute != null)
                                {
                                    foreach (var value in pilotAttribute.XPathSelectElements("Values/child::node()"))
                                    {
                                        var attributeValue = new AttributeValueChange();
                                        attributeValue.NewValue = (string)value;
                                        attributeValue.ValueModificationType = this.TestAttributeValueChange(attributeName, attributeValue.NewValue, DataRowState.Added) ? DataRowState.Added : DataRowState.Unchanged;
                                        if (attributeChange.HasReference)
                                        {
                                            this.TryDereferenceValue(attributeValue);
                                        }

                                        attributeChange.AttributeValues.Add(attributeValue);
                                    }
                                }

                                if (productionAttribute != null)
                                {
                                    foreach (var value in productionAttribute.XPathSelectElements("Values/child::node()"))
                                    {
                                        if (this.TestAttributeValueChange(attributeName, (string)value, DataRowState.Deleted))
                                        {
                                            var attributeValue = new AttributeValueChange();
                                            attributeValue.OldValue = (string)value;
                                            attributeValue.ValueModificationType = DataRowState.Deleted;
                                            if (attributeChange.HasReference)
                                            {
                                                this.TryDereferenceValue(attributeValue);
                                            }

                                            attributeValue.NewValue = attributeValue.OldValue;
                                            attributeChange.AttributeValues.Add(attributeValue);
                                        }
                                    }
                                }
                            }

                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    var attributeValue = new AttributeValueChange();
                    switch (this.CurrentChangeObjectState)
                    {
                        case "Create":
                            {
                                attributeChange.AttributeModificationType = DataRowState.Added;
                                attributeValue.ValueModificationType = attributeChange.AttributeModificationType;
                                attributeValue.NewValue = (string)pilotAttribute.Element("Value");
                                if (attributeChange.HasReference)
                                {
                                    this.TryDereferenceValue(attributeValue);
                                }
                            }

                            break;
                        case "Delete":
                            {
                                attributeChange.AttributeModificationType = DataRowState.Deleted;
                                attributeValue.ValueModificationType = attributeChange.AttributeModificationType;
                                attributeValue.OldValue = (string)productionAttribute.Element("Value");
                                attributeValue.NewValue = attributeValue.OldValue;
                                if (attributeChange.HasReference)
                                {
                                    this.TryDereferenceValue(attributeValue);
                                }
                            }

                            break;
                        case "Put":
                            {
                                attributeChange.AttributeModificationType = this.CurrentChangeObject.XPathSelectElement(ServiceCommonDocumenter.GetChangeAttributeXPath(attributeName, DataRowState.Modified)) == null ? DataRowState.Unchanged : DataRowState.Modified;
                                attributeValue.ValueModificationType = attributeChange.AttributeModificationType;
                                attributeValue.NewValue = pilotAttribute != null ? (string)pilotAttribute.Element("Value") : string.Empty;
                                attributeValue.OldValue = productionAttribute != null ? (string)productionAttribute.Element("Value") : string.Empty;
                                if (attributeChange.HasReference)
                                {
                                    this.TryDereferenceValue(attributeValue);
                                }
                            }

                            break;
                    }

                    attributeChange.AttributeValues.Add(attributeValue);
                }

                return attributeChange;
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Attribute Name : '{0}'", attributeName);
            }
        }

        /// <summary>
        /// Tries to dereference the attribute value if it contains a reference.
        /// </summary>
        /// <param name="attributeValue">The attribute value.</param>
        protected void TryDereferenceValue(AttributeValueChange attributeValue)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (attributeValue == null)
                {
                    return;
                }

                attributeValue.NewId = attributeValue.NewValue;
                attributeValue.OldId = attributeValue.OldValue;

                if (!string.IsNullOrEmpty(attributeValue.NewId))
                {
                    var resolvedValue = this.TryResolveReferences(attributeValue.NewId, true, attributeValue.ValueModificationType);
                    attributeValue.NewValue = resolvedValue.Item1;
                    attributeValue.NewValueText = resolvedValue.Item2;
                }

                if (!string.IsNullOrEmpty(attributeValue.OldId))
                {
                    var resolvedValue = this.TryResolveReferences(attributeValue.OldId, false, attributeValue.ValueModificationType);
                    attributeValue.OldValue = resolvedValue.Item1;
                    attributeValue.OldValueText = resolvedValue.Item2;

                    if (attributeValue.ValueModificationType == DataRowState.Deleted)
                    {
                        attributeValue.NewValue = attributeValue.OldValue;
                        attributeValue.NewId = attributeValue.OldId;
                        attributeValue.NewValueText = attributeValue.OldValueText;
                    }
                }
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the xpath for locating the specified attribute of the current change object in the specified config.
        /// </summary>
        /// <param name="input">The input string containing of the potential references.</param>
        /// <param name="pilotConfig">if set to <c>true</c>, indicates that this is a pilot configuration. Otherwise, this is a production configuration.</param>
        /// <param name="valueModificationType"> The value modification type.</param>
        /// <returns>The a tuple of resolved markup and plain text value.</returns>
        protected Tuple<string, string> TryResolveReferences(string input, bool pilotConfig, DataRowState valueModificationType)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var guids = ExtractGuid(input);
                var resolvedReferences = new Dictionary<string, Tuple<string, string>>();
                var resolvedInput = input;

                foreach (var guid in guids.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    var objectId = "urn:uuid:" + guid;
                    var displayNameMarkup = guid;
                    var displayName = input;

                    var configXml = pilotConfig ? this.PilotXml : this.ProductionXml;
                    var referencedObject = configXml.XPathSelectElement(ServiceCommonDocumenter.GetObjectXPath(objectId, pilotConfig));
                    if (referencedObject != null)
                    {
                        displayName = (string)referencedObject.XPathSelectElement(ServiceCommonDocumenter.GetRelativeAttributeXPath("DisplayName") + "/Value");

                        if (!string.IsNullOrEmpty(displayName))
                        {
                            displayNameMarkup = ServiceCommonDocumenter.GetJumpToBookmarkLocationMarkup(displayName, objectId, valueModificationType);
                        }

                        if (input == objectId)
                        {
                            input = displayName;
                        }
                    }

                    resolvedReferences.Add(guid, new Tuple<string, string>(displayNameMarkup, displayName));
                }

                foreach (string guid in resolvedReferences.Keys)
                {
                    resolvedInput = resolvedInput.Replace("urn:uuid:" + guid, resolvedReferences[guid].Item1).Replace(guid, resolvedReferences[guid].Item1);
                }

                return new Tuple<string, string>(resolvedInput, input);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Gets the Attribute Change for the localization configuration of the given attribute.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        /// <param name="locale">The locale of the localization configuration.</param>
        /// <returns>The localization configuration change.</returns>
        protected AttributeChange GetAttributeLocalizationChange(string attributeName, string locale)
        {
            var attributeChange = new AttributeChange(attributeName, locale);
            var attributeValue = new AttributeValueChange();
            switch (this.CurrentChangeObjectState)
            {
                case "Create":
                    {
                        attributeChange.AttributeModificationType = DataRowState.Added;
                        attributeValue.NewValue = (string)this.PilotXml.XPathSelectElement(this.GetAttributeXPath(attributeName, true, locale) + "/Value");
                    }

                    break;
                case "Delete":
                    {
                        attributeChange.AttributeModificationType = DataRowState.Deleted;
                        attributeValue.OldValue = (string)this.ProductionXml.XPathSelectElement(this.GetAttributeXPath(attributeName, false, locale) + "/Value");
                    }

                    break;
                case "Put":
                    {
                        attributeChange.AttributeModificationType = this.CurrentChangeObject.XPathSelectElement(ServiceCommonDocumenter.GetChangeAttributeXPath(attributeName, DataRowState.Modified, locale)) == null ? DataRowState.Unchanged : DataRowState.Modified;
                        attributeValue.NewValue = (string)this.PilotXml.XPathSelectElement(this.GetAttributeXPath(attributeName, true) + "/Value");
                        attributeValue.OldValue = (string)this.ProductionXml.XPathSelectElement(this.GetAttributeXPath(attributeName, false) + "/Value");
                    }

                    break;
                default:
                    break;
            }

            attributeChange.AttributeValues.Add(attributeValue);

            return attributeChange;
        }

        /// <summary>
        /// Writes the section header
        /// </summary>
        /// <param name="title">The section title</param>
        /// <param name="level">The section header level</param>
        /// <param name="environment">The config environment.</param>
        protected void WriteSectionHeader(string title, int level, ConfigEnvironment environment)
        {
            this.Environment = environment; // reset the config environment
            this.WriteSectionHeader(title, level);
        }

        /// <summary>
        /// Gets the CSS visibility class. For FIM Service always return NoHide (for now).
        /// </summary>
        /// <returns>The CSS visibility class, either CanHide or empty string</returns>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification = "Reviewed.")]
        protected override string GetCssVisibilityClass()
        {
            return Documenter.NoHide;
        }

        #region Simple Multivalue Values Diffgram DataSet

        /// <summary>
        /// Creates simple multi-value values diffgram dataset with two columns one for Value text and other for Value markup.
        /// The setting name is expected to be contained in the table header. e.g Set - Criteria-based Members and Manually-managed Members
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateSimpleMultivalueValuesDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("SimpleMultivalueValues") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("ValueText");
                var column2 = new DataColumn("ValueMarkup");

                table.Columns.Add(column1);
                table.Columns.Add(column2);

                var printTable = Documenter.GetPrintTable();

                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                var diffgramTable = Documenter.CreateDiffgramTable(table);

                this.DiffgramDataSet = new DataSet("SimpleMultivalueValues") { Locale = CultureInfo.InvariantCulture };
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
        /// Fills the current object.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        protected void FillSimpleMultivalueValuesDiffgramDataSet(string attributeName)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                this.AddSimpleMultivalueValuesRows(attributeName);
                this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Adds rows to the simple multi-value values diffgram for the given attribute.
        /// </summary>
        /// <param name="attributeName">The name of the attribute.</param>
        protected void AddSimpleMultivalueValuesRows(string attributeName)
        {
            var diffgramTable = this.DiffgramDataSet.Tables[0];
            var attributeChange = this.GetAttributeChange(attributeName);
            foreach (var attributeValue in attributeChange.AttributeValues)
            {
                Documenter.AddRow(diffgramTable, new object[] { attributeValue.NewValueText, attributeValue.NewValue, attributeValue.ValueModificationType, attributeValue.OldValueText, attributeValue.OldValue });
            }
        }

        #endregion Simple Multivalue Values Diffgram DataSet

        #region Simple Multivalue Ordered Settings Diffgram DataSet

        /// <summary>
        /// Creates simple multi-value ordered settings diffgram dataset.
        /// The Parent table contains the configuration name and the child table contains one or more values for that setting.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateSimpleMultivalueOrderedSettingsDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("SimpleOrderedSettings") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("DisplayOrderIndex", typeof(int));
                var column2 = new DataColumn("Setting");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("SimpleOrderedSettingsConfiguration") { Locale = CultureInfo.InvariantCulture };
                var column12 = new DataColumn("DisplayOrderIndex", typeof(int));
                var column22 = new DataColumn("ConfigurationIndex", typeof(int)); // needed for multivalue attributes
                var column32 = new DataColumn("Configuration"); // needed for sorting multivalues
                var column42 = new DataColumn("ConfigurationMarkup");
                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.Columns.Add(column32);
                table2.Columns.Add(column42);
                table2.PrimaryKey = new[] { column12, column22 };

                var printTable = Documenter.GetPrintTable();

                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 2 }, { "Hidden", true }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 3 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                var diffgramTable = Documenter.CreateDiffgramTable(table);
                var diffgramTable2 = Documenter.CreateDiffgramTable(table2);

                this.DiffgramDataSet = new DataSet("SimpleOrderedSettings") { Locale = CultureInfo.InvariantCulture };
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
        /// Fills the current object.
        /// </summary>
        /// <param name="settings">The array of key-value pairs of setting display name and corresponding attribute name.</param>
        protected void FillSimpleMultivalueOrderedSettingsDiffgramDataSet(KeyValuePair<string, string>[] settings)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (settings != null)
                {
                    foreach (var setting in settings)
                    {
                        var attributeChange = this.GetAttributeChange(setting.Value);
                        this.AddSimpleMultivalueOrderedRows(setting.Key, attributeChange);
                    }
                }

                this.DiffgramDataSet = Documenter.SortDataSet(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Adds rows to the simple multi-value ordered diffgram for the given attribute.
        /// </summary>
        /// <param name="setting">The name of the configuration.</param>
        /// <param name="attributeChange">The attribute change.</param>
        protected void AddSimpleMultivalueOrderedRows(string setting, AttributeChange attributeChange)
        {
            var diffgramTable = this.DiffgramDataSet.Tables[0];
            var diffgramTable2 = this.DiffgramDataSet.Tables[1];

            var rowIndex = diffgramTable.Rows.Count;
            Documenter.AddRow(diffgramTable, new object[] { rowIndex, setting, DataRowState.Unchanged });

            if (attributeChange != null)
            {
                for (var attributeValueIndex = 0; attributeValueIndex < attributeChange.AttributeValues.Count; ++attributeValueIndex)
                {
                    Documenter.AddRow(diffgramTable2, new object[] { rowIndex, attributeValueIndex, attributeChange.AttributeValues[attributeValueIndex].NewValueText, attributeChange.AttributeValues[attributeValueIndex].NewValue, attributeChange.AttributeValues[attributeValueIndex].ValueModificationType, attributeChange.AttributeValues[attributeValueIndex].OldValueText, attributeChange.AttributeValues[attributeValueIndex].OldValue });
                }
            }
        }

        #endregion Simple Multivalue Ordered Settings Diffgram DataSet

        #region Nested Multivalue Ordered Settings Diffgram DataSet

        /// <summary>
        /// Creates a nested multi-value ordered settings diffgram dataset.
        /// The Parent table contains the configuration name.
        /// The child table contains the sub-setting.
        /// The grand child table contains one or more values for that sub-setting.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateNestedMultivalueOrderedSettingsDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("Sections") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Section");

                table.Columns.Add(column1);
                table.PrimaryKey = new[] { column1 };

                var table2 = new DataTable("Sub-sections") { Locale = CultureInfo.InvariantCulture };

                var column12 = new DataColumn("Section");
                var column22 = new DataColumn("Sub-section");

                table2.Columns.Add(column12);
                table2.Columns.Add(column22);
                table2.PrimaryKey = new[] { column12, column22 };

                var table3 = new DataTable("Settings") { Locale = CultureInfo.InvariantCulture };

                var column13 = new DataColumn("Section");
                var column23 = new DataColumn("Sub-section");
                var column33 = new DataColumn("ConfigurationIndex", typeof(int)); // needed for multivalue attributes
                var column43 = new DataColumn("Configuration"); // needed for sorting multivalues
                var column53 = new DataColumn("ConfigurationMarkup");

                table3.Columns.Add(column13);
                table3.Columns.Add(column23);
                table3.Columns.Add(column33);
                table3.Columns.Add(column43);
                table3.Columns.Add(column53);
                table3.PrimaryKey = new[] { column13, column23, column33 };

                var printTable = Documenter.GetPrintTable();

                // Table 1
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 2
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 1 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                // Table 3
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 3 }, { "Hidden", true }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 2 }, { "ColumnIndex", 4 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                var diffgramTable = Documenter.CreateDiffgramTable(table);
                var diffgramTable2 = Documenter.CreateDiffgramTable(table2);
                var diffgramTable3 = Documenter.CreateDiffgramTable(table3);

                this.DiffgramDataSet = new DataSet("Sections") { Locale = CultureInfo.InvariantCulture };
                this.DiffgramDataSet.Tables.Add(diffgramTable);
                this.DiffgramDataSet.Tables.Add(diffgramTable2);
                this.DiffgramDataSet.Tables.Add(diffgramTable3);
                this.DiffgramDataSet.Tables.Add(printTable);

                // set up data relations
                var dataRelation12 = new DataRelation("DataRelation12", new[] { diffgramTable.Columns[0] }, new[] { diffgramTable2.Columns[0] }, false);
                this.DiffgramDataSet.Relations.Add(dataRelation12);

                var dataRelation23 = new DataRelation("DataRelation23", new[] { diffgramTable2.Columns[0], diffgramTable2.Columns[1] }, new[] { diffgramTable3.Columns[0], diffgramTable3.Columns[1] }, false);
                this.DiffgramDataSet.Relations.Add(dataRelation23);

                Documenter.AddRowVisibilityStatusColumn(this.DiffgramDataSet);

                this.DiffgramDataSet.AcceptChanges();
                this.DiffgramDataSets.Add(this.DiffgramDataSet);
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Nested Multivalue Ordered Settings Diffgram DataSet

        #region Simple Summary Settings Diffgram DataSet

        /// <summary>
        /// Creates a simple summary settings diffgram dataset.
        /// First two columns are DisplayName (hidden, sort order) and DisplayNameMarkup
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateSimpleSummarySettingsDiffgramDataSet(int columnCount)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("SimpleSummarySettings") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Display Name"); // for sorting purposes
                var column2 = new DataColumn("Display Name Markup");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                for (var i = 2; i < columnCount; ++i)
                {
                    table.Columns.Add(new DataColumn("Column" + (i + 1)));
                }

                table.PrimaryKey = new[] { column1 };

                var printTable = Documenter.GetPrintTable();
                for (var i = 0; i < columnCount; ++i)
                {
                    printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", i }, { "Hidden", i == 0 }, { "SortOrder", i == 0 ? 1 : -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                }

                var diffgramTable = Documenter.CreateDiffgramTable(table);

                this.DiffgramDataSet = new DataSet("SimpleSummarySettings") { Locale = CultureInfo.InvariantCulture };
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

        #endregion Simple Summary Settings Diffgram DataSet

        #region Simple Ordered Settings Diffgram DataSet

        /// <summary>
        /// Creates the simple ordered settings diffgram data sets.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        protected void CreateSimpleOrderedSettingsDiffgramDataSet(int columnCount)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'.", columnCount);

            try
            {
                this.CreateSimpleOrderedSettingsDiffgramDataSet(columnCount, 2, false);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'.", columnCount);
            }
        }

        /// <summary>
        /// Creates the simple ordered settings diffgram data sets.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        /// <param name="alphabeticOrder">if set to <c>true</c>, the rows are sorted alphabetically.</param>
        protected void CreateSimpleOrderedSettingsDiffgramDataSet(int columnCount, bool alphabeticOrder)
        {
            Logger.Instance.WriteMethodEntry("Column Count: '{0}'.", columnCount);

            try
            {
                this.CreateSimpleOrderedSettingsDiffgramDataSet(columnCount, 2, alphabeticOrder);
            }
            finally
            {
                Logger.Instance.WriteMethodExit("Column Count: '{0}'.", columnCount);
            }
        }

        /// <summary>
        /// Creates a simple ordered settings diffgram dataset.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        /// <param name="keyCount">The key count.</param>
        /// <param name="alphabeticOrder">if set to <c>true</c>, the rows are sorted alphabetically.</param>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateSimpleOrderedSettingsDiffgramDataSet(int columnCount, int keyCount, bool alphabeticOrder)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("SimpleOrderedSettings") { Locale = CultureInfo.InvariantCulture };

                if (alphabeticOrder)
                {
                    table.Columns.Add(new DataColumn("Column1", typeof(string)));
                }
                else
                {
                    table.Columns.Add(new DataColumn("Column1", typeof(int)));
                }

                for (var i = 1; i < columnCount; ++i)
                {
                    table.Columns.Add(new DataColumn("Column" + (i + 1)));
                }

                var primaryKey = new List<DataColumn>(keyCount);
                for (var i = 0; i < keyCount; ++i)
                {
                    primaryKey.Add(table.Columns[i]);
                }

                table.PrimaryKey = primaryKey.ToArray();

                var printTable = this.GetSimpleOrderedSettingsPrintTable(columnCount);

                var diffgramTable = Documenter.CreateDiffgramTable(table);

                this.DiffgramDataSet = new DataSet("SimpleOrderedSettings") { Locale = CultureInfo.InvariantCulture };
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

        #endregion Simple Ordered Settings Diffgram DataSet

        /// <summary>
        /// Prints a simple section header based on the current objects display name and object guid
        /// </summary>
        /// <param name="level">The section level</param>
        protected void PrintSimpleSectionHeader(int level)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var sectionTitle = this.GetAttributeChange("DisplayName").NewValue;
                var sectionGuid = this.GetAttributeChange("ObjectID").NewValue;

                Logger.Instance.WriteInfo("Processing " + this.CurrentChangeObjectType + ":  " + sectionTitle);

                this.WriteSectionHeader(sectionTitle, level, sectionGuid);
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints simple settings section table
        /// </summary>
        /// <param name="columnNames">The table header column names</param>
        protected void PrintSimpleSettingsSectionTable(OrderedDictionary columnNames)
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = columnNames == null || columnNames.Count == 0 ? null : Documenter.GetSimpleSettingsHeaderTable(columnNames);
                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #region Localization Configurartion

        /// <summary>
        /// Processes the localization configuration of the object.
        /// </summary>
        protected void ProcessLocalizationConfiguration()
        {
            this.CreateLocalizationConfigurationDiffgramDataSet();

            this.FillLocalizationConfigurationDiffgramDataSet();

            this.PrintLocalizationConfiguration();
        }

        /// <summary>
        /// Creates localization configuration diffgram dataset.
        /// </summary>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No good reason to call Dispose() on DataTable and DataColumn.")]
        protected void CreateLocalizationConfigurationDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var table = new DataTable("LocalizationConfiguration") { Locale = CultureInfo.InvariantCulture };

                var column1 = new DataColumn("Locale");
                var column2 = new DataColumn("AttributeName");
                var column3 = new DataColumn("AttributeValue");

                table.Columns.Add(column1);
                table.Columns.Add(column2);
                table.Columns.Add(column3);
                table.PrimaryKey = new[] { column1, column2 };

                var printTable = Documenter.GetPrintTable();

                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 0 }, { "Hidden", false }, { "SortOrder", 1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 1 }, { "Hidden", false }, { "SortOrder", 2 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());
                printTable.Rows.Add((new OrderedDictionary { { "TableIndex", 0 }, { "ColumnIndex", 2 }, { "Hidden", false }, { "SortOrder", -1 }, { "BookmarkIndex", -1 }, { "JumpToBookmarkIndex", -1 }, { "ChangeIgnored", false } }).Values.Cast<object>().ToArray());

                var diffgramTable = Documenter.CreateDiffgramTable(table);

                this.DiffgramDataSet = new DataSet("LocalizationConfiguration") { Locale = CultureInfo.InvariantCulture };
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
        /// Fills localization configuration diffgram dataset.
        /// </summary>
        protected void FillLocalizationConfigurationDiffgramDataSet()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                var diffgramTable = this.DiffgramDataSet.Tables[0];

                var objectIdentifier = (string)this.CurrentChangeObject.Element("SourceObjectIdentifier");
                var localizedAttributes = this.PilotXml.XPathSelectElements(ServiceCommonDocumenter.PilotExportObjectRootXPath + "/ResourceManagementObject[ObjectIdentifier = '" + objectIdentifier + "']/LocalizedResourceManagementAttributes/LocalizedResourceManagementAttribute");

                // sort by locale
                localizedAttributes = from localizedAttribute in localizedAttributes
                                      let locale = (string)localizedAttribute.Element("Culture")
                                      orderby locale
                                      select localizedAttribute;

                foreach (var localizedAttribute in localizedAttributes)
                {
                    var locale = (string)localizedAttribute.Element("Culture");
                    var attributes = localizedAttribute.XPathSelectElements(".//ResourceManagementAttribute");

                    // sort by attribute name
                    attributes = from attribute in attributes
                                 let attributeName = (string)attribute.Element("AttributeName")
                                 orderby attributeName
                                 select attribute;

                    foreach (var attribute in attributes)
                    {
                        var attributeName = (string)attribute.Element("AttributeName");
                        var attributeChange = this.GetAttributeLocalizationChange(attributeName, locale);
                        Documenter.AddRow(diffgramTable, new object[] { locale, attributeName, attributeChange.AttributeValues[0].NewValue, attributeChange.AttributeModificationType, attributeChange.AttributeValues[0].OldValue });
                    }
                }

                this.DiffgramDataSet.AcceptChanges();
            }
            finally
            {
                Logger.Instance.WriteMethodExit();
            }
        }

        /// <summary>
        /// Prints the localization configuration.
        /// </summary>
        protected void PrintLocalizationConfiguration()
        {
            Logger.Instance.WriteMethodEntry();

            try
            {
                if (this.DiffgramDataSet.Tables[0].Rows.Count != 0)
                {
                    var headerTable = Documenter.GetSimpleSettingsHeaderTable("Localization", new OrderedDictionary { { "Locale", 20 }, { "Attribute", 30 }, { "Value", 50 } });

                    this.WriteTable(this.DiffgramDataSet.Tables[0], headerTable);
                }
            }
            finally
            {
                this.ResetDiffgram(); // reset the diffgram variables
                Logger.Instance.WriteMethodExit();
            }
        }

        #endregion Localization Configurartion

        #region Helper Classes

        /// <summary>
        /// Represents the change to the attribute value.
        /// </summary>
        protected class AttributeValueChange
        {
            /// <summary>
            /// The old value text.
            /// </summary>
            private string oldValueText;

            /// <summary>
            /// The new value text.
            /// </summary>
            private string newValueText;

            /// <summary>
            /// Gets or sets the old attribute value.
            /// If it's a reference attribute, then the jump link markup for the display name of the referenced attribute.
            /// </summary>
            public string OldValue { get; set; }

            /// <summary>
            /// Gets or sets the new attribute value.
            /// If it's a reference attribute, then the jump link markup for the display name of the referenced attribute.
            /// </summary>
            public string NewValue { get; set; }

            /// <summary>
            /// Gets or sets the old ObjectID of the attribute if it's a reference attribute.
            /// </summary>
            public string OldId { get; set; }

            /// <summary>
            /// Gets or sets the new ObjectID of the attribute if it's a reference attribute.
            /// </summary>
            public string NewId { get; set; }

            /// <summary>
            /// Gets or sets the old attribute value in plain text. 
            /// Meant to be used to get the raw display name without any markup if it's a reference attribute.
            /// </summary>
            public string OldValueText
            {
                get
                {
                    return !string.IsNullOrEmpty(this.oldValueText) ? this.oldValueText : this.OldValue;
                }

                set
                {
                    this.oldValueText = value;
                }
            }

            /// <summary>
            /// Gets or sets the new attribute value in plain text.
            /// Meant to be used to get the raw display name without any markup if it's a reference attribute.
            /// </summary>
            public string NewValueText
            {
                get
                {
                    return !string.IsNullOrEmpty(this.newValueText) ? this.newValueText : this.NewValue;
                }

                set
                {
                    this.newValueText = value;
                }
            }

            /// <summary>
            /// Gets or sets the type of modification to the attribute value.
            /// </summary>
            public DataRowState ValueModificationType { get; set; }
        }

        /// <summary>
        /// Represents the change to the attribute.
        /// </summary>
        protected class AttributeChange
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="AttributeChange"/> class.
            /// </summary>
            /// <param name="attributeName">The name of the attribute.</param>
            public AttributeChange(string attributeName) : this(attributeName, ServiceCommonDocumenter.InvariantLocale)
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="AttributeChange"/> class.
            /// </summary>
            /// <param name="attributeName">The name of the attribute.</param>
            /// <param name="locale">The locale of the localization configuration</param>
            public AttributeChange(string attributeName, string locale)
            {
                this.AttributeName = attributeName;
                this.AttributeModificationType = DataRowState.Unchanged;
                this.AttributeValues = new List<AttributeValueChange>();
                this.Locale = locale;
                this.IsMultivalue = false;
                this.HasReference = false;
            }

            /// <summary>
            /// Gets or sets the name of the attribute.
            /// </summary>
            public string AttributeName { get; set; }

            /// <summary>
            /// Gets or sets the type of modification to the attribute value.
            /// </summary>
            public DataRowState AttributeModificationType { get; set; }

            /// <summary>
            /// Gets or sets the list of attribute values.
            /// </summary>
            public List<AttributeValueChange> AttributeValues { get; set; }

            /// <summary>
            /// Gets the old value of the attribute, it it's a single value attribute.
            /// Otherwise the first item in the old multi-value attribute list.
            /// </summary>
            public string OldValue
            {
                get
                {
                    if (this.AttributeValues.Count == 0)
                    {
                        return string.Empty;
                    }
                    else
                    {
                        return this.AttributeValues[0].OldValue;
                    }
                }
            }

            /// <summary>
            /// Gets the new value of the attribute, it it's a single value attribute.
            /// Otherwise the first item in the new multi-value attribute list.
            /// </summary>
            public string NewValue
            {
                get
                {
                    if (this.AttributeValues.Count == 0)
                    {
                        return string.Empty;
                    }
                    else
                    {
                        return this.AttributeValues[0].NewValue;
                    }
                }
            }

            /// <summary>
            /// Gets the old ObjectID of the attribute if it's a reference attribute.
            /// Otherwise the first item in the old multi-value reference attribute list.
            /// </summary>
            public string OldId
            {
                get
                {
                    if (this.AttributeValues.Count == 0)
                    {
                        return string.Empty;
                    }
                    else
                    {
                        return this.AttributeValues[0].OldId;
                    }
                }
            }

            /// <summary>
            /// Gets the new ObjectID of the attribute if it's a reference attribute.
            /// Otherwise the first item in the new multi-value reference attribute list.
            /// </summary>
            public string NewId
            {
                get
                {
                    if (this.AttributeValues.Count == 0)
                    {
                        return string.Empty;
                    }
                    else
                    {
                        return this.AttributeValues[0].NewId;
                    }
                }
            }

            /// <summary>
            /// Gets the old raw value of the attribute or the the first item in the old multi-value reference attribute list.
            /// </summary>
            public string OldValueText
            {
                get
                {
                    if (this.AttributeValues.Count == 0)
                    {
                        return string.Empty;
                    }
                    else
                    {
                        return this.AttributeValues[0].OldValueText;
                    }
                }
            }

            /// <summary>
            /// Gets the new raw value of the attribute or the the first item in the old multi-value reference attribute list.
            /// </summary>
            public string NewValueText
            {
                get
                {
                    if (this.AttributeValues.Count == 0)
                    {
                        return string.Empty;
                    }
                    else
                    {
                        return this.AttributeValues[0].NewValueText;
                    }
                }
            }

            /// <summary>
            /// Gets or sets the locale of the attribute localization configuration
            /// </summary>
            public string Locale { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the attribute is multi-value
            /// </summary>
            public bool IsMultivalue { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the attribute is reference
            /// </summary>
            public bool HasReference { get; set; }
        }
    }

    #endregion Helper Classes
}
