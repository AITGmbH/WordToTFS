using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.CompilerServices;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.TfsHelper;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Service.Configuration.Serialization;
using AIT.TFS.SyncService.Service.WorkItemObjects;
using System.Linq;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// Class that represents a work item.
    /// </summary>

    public class ConfigurationItem : IConfigurationItem
    {
        #region Private fields

        private readonly IDictionary<string, IConverter> _converters;
        private readonly string _stateTransitionField;
        private readonly Dictionary<string, string> _stateTransitions;
        private readonly IConfiguration _configuration;
        private readonly string _imageFilename;
        private Image _image;

        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationItem"/> class.
        /// </summary>
        /// <param name="configuration">Configuration to copy the settings from.</param>
        private ConfigurationItem(ConfigurationItem configuration)
        {
            WorkItemType = configuration.WorkItemType;
            WorkItemSubtypeField = configuration.WorkItemSubtypeField;
            WorkItemSubtypeValue = configuration.WorkItemSubtypeValue;
            WorkItemTypeMapping = configuration.WorkItemTypeMapping;
            ReqTableIdentifierExpression = configuration.ReqTableIdentifierExpression;
            ReqTableCellCol = configuration.ReqTableCellCol;
            ReqTableCellRow = configuration.ReqTableCellRow;
            RelatedTemplate = configuration.RelatedTemplate;
            RelatedTemplateFile = configuration.RelatedTemplateFile;
            HideElementInWord = configuration.HideElementInWord;
            _image = configuration._image;
            Level = configuration.Level;
            FieldConfigurations = new List<IConfigurationFieldItem>(configuration.FieldConfigurations);

            _stateTransitionField = configuration._stateTransitionField;
            if (configuration._stateTransitions != null)
                _stateTransitions = new Dictionary<string, string>(configuration._stateTransitions);
            _converters = new Dictionary<string, IConverter>(configuration._converters);

            Links = new List<IConfigurationLinkItem>(configuration.Links);
            FieldToLinkedItemConfiguration = new List<IConfigurationFieldToLinkedItem>(configuration.FieldToLinkedItemConfiguration);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationItem"/> class.
        /// </summary>
        /// <param name="configuration">Configuration, use to get access the the absolute path.</param>
        /// <param name="mapping">Mapping used to retrieve values.</param>
        /// <param name="templateSubFolder">Folder in which to search for referenced files like work item templates and icons</param>
        public ConfigurationItem(IConfiguration configuration, MappingElement mapping, string templateSubFolder)
        {
            SyncServiceTrace.I(Properties.Resources.LogService_LoadWorkItem, mapping.WorkItemType);

            // Set the simple properties
            WorkItemType = mapping.WorkItemType;
            WorkItemSubtypeField = mapping.WorkItemSubtypeField;
            WorkItemSubtypeValue = mapping.WorkItemSubtypeValue;
            WorkItemTypeMapping = mapping.MappingWorkItemType;
            ReqTableIdentifierExpression = mapping.AssignRegularExpression;
            ReqTableCellRow = mapping.AssignCellRow;
            ReqTableCellCol = mapping.AssignCellCol;
            RelatedTemplate = mapping.RelatedTemplate;
            _configuration = configuration;
            RelatedTemplateFile = Path.Combine(templateSubFolder, RelatedTemplate);
            HideElementInWord = mapping.HideElementInWord;
            PreOperations = null;
            PostOperations = null;

            _imageFilename = Path.Combine(templateSubFolder, mapping.ImageFile);

            FieldConfigurations = new List<IConfigurationFieldItem>();

            if (mapping.PreOperations != null && mapping.PreOperations.Count > 0)
            {
                PreOperations = new List<IConfigurationTestOperation>();
                foreach (var operation in mapping.PreOperations)
                    PreOperations.Add(new ConfigurationTestOperation(operation));
            }
            if (mapping.PostOperations != null && mapping.PostOperations.Count > 0)
            {
                PostOperations = new List<IConfigurationTestOperation>();
                foreach (var operation in mapping.PostOperations)
                    PostOperations.Add(new ConfigurationTestOperation(operation));
            }

            // Insert a field for the invisible stack rank
            if (_configuration.UseStackRank)
            {
                var stackRankConfig = FieldConfigurations.FirstOrDefault(x => x.ReferenceFieldName.Equals("Microsoft.VSTS.Common.StackRank", StringComparison.OrdinalIgnoreCase));
                if (stackRankConfig != null)
                {
                    SyncServiceTrace.W("Mapping the stack rank field is not supported in combination with UseStackRank=true");
                    FieldConfigurations.Remove(stackRankConfig);
                }
                FieldConfigurations.Add(
                    new ConfigurationFieldItem("Microsoft.VSTS.Common.StackRank", null, FieldValueType.PlainText,
                                                Direction.OtherToTfs, 0, 0, null, false, HandleAsDocumentType.All, null, string.Empty, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, String.Empty));
            }

            // Read the configured state transitions if some are defined
            if (mapping.StateTransitions != null)
            {
                _stateTransitionField = mapping.StateTransitions.FieldName;
                _stateTransitions = new Dictionary<string, string>();

                foreach (var transition in mapping.StateTransitions.Items)
                {
                    _stateTransitions.Add(transition.From, transition.To);
                }
            }

            foreach (MappingField field in mapping.Fields)
            {
                // This fixes some items having their ids configured as html.
                if (field.Name.Equals(FieldReferenceNames.SystemId)) field.FieldValueType = FieldValueType.PlainText;

                ConfigurationFieldItemDefaultValue cfidv = null;
                if (field.DefaultValue != null)
                    cfidv = new ConfigurationFieldItemDefaultValue(field.DefaultValue.ShowName,
                                                                   field.DefaultValue.DefaultValue);
                FieldConfigurations.Add(new ConfigurationFieldItem(field.Name, field.MappingName, field.FieldValueType,
                                                                   field.Direction, field.TableRow, field.TableCol,
                                                                   field.TestCaseStepDelimiter, field.HandleAsDocument,
                                                                   field.HandleAsDocumentMode, field.OLEMarkerField,
                                                                   field.OLEMarkerValue, cfidv,
                                                                   field.ShapeOnlyWorkaroundMode,
                                                                   field.TableCol != 0 && field.TableRow != 0,
                                                                   field.WordBookmark,
                                                                   field.VariableName,
                                                                   field.DateTimeFormat
                                                                   ));
            }

            // If the state transition field is not mapped, add it as unmapped field (so it gets queried anyways)
            var stateTransition = FieldConfigurations.FirstOrDefault(x => x.ReferenceFieldName.Equals(_stateTransitionField));
            if (stateTransition == null && _stateTransitionField != null)
            {
                FieldConfigurations.Add(new ConfigurationFieldItem(_stateTransitionField, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 0, 0, string.Empty, false, HandleAsDocumentType.OleOnDemand, null, string.Empty, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));
            }

            // Set the converters
            _converters = new Dictionary<string, IConverter>();

            SyncServiceTrace.I(Properties.Resources.LogService_LoadWorkItem_Converters, mapping.Converters.Length, mapping.WorkItemType);

            foreach (MappingConverter converter in mapping.Converters)
            {
                var valueMapper = new SimpleValueMapper(converter.FieldName);
                foreach (MappingConverterValue value in converter.Values)
                    valueMapper.AddMapping(value.Text, value.MappingText);
                _converters.Add(valueMapper.FieldName.ToUpperInvariant(), valueMapper);
            }

            Links = new List<IConfigurationLinkItem>();
            if (mapping.Links != null)
            {
                foreach (MappingLink link in mapping.Links)
                {
                    var cli = new ConfigurationLinkItem(link.LinkValueType, link.LinkedWorkItemTypes, link.Direction, link.TableRow, link.TableCol, link.Overwrite, link.LinkSeparator, link.LinkFormat, link.AutomaticLinkWorkItemType, link.AutomaticLinkWorkItemSubtypeField, link.AutomaticLinkWorkItemSubtypeValue, link.AutomaticLinkSuppressWarnings);
                    Links.Add(cli);
                }
            }

            // Read the 'field to linked work item' configuration.
            FieldToLinkedItemConfiguration = new List<IConfigurationFieldToLinkedItem>();
            if (mapping.FieldsToLinkedItems != null)
            {
                foreach (var fieldToLinkedItem in mapping.FieldsToLinkedItems)
                {
                    FieldToLinkedItemConfiguration.Add(new ConfigurationFieldToLinkedItem(fieldToLinkedItem));
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationItem"/> class.
        /// </summary>
        /// <param name="configuration">Configuration, used to get the absolute path.</param>
        /// <param name="header">Mapping used to retrieve values.</param>
        /// <param name="templateSubFolder">Folder in which to search for referenced files like work item templates and icons</param>
        public ConfigurationItem(IConfiguration configuration, MappingHeader header, string templateSubFolder)
        {
            SyncServiceTrace.I("Loading header {0}", header.Identifier);

            // Set the simple properties
            WorkItemType = header.Identifier;
            WorkItemSubtypeField = null;
            WorkItemSubtypeValue = null;
            WorkItemTypeMapping = header.Identifier;
            ReqTableIdentifierExpression = header.Identifier + " ";
            ReqTableCellRow = header.Row;
            ReqTableCellCol = header.Column;
            RelatedTemplate = header.RelatedTemplate;

            _configuration = configuration;
            _imageFilename = Path.Combine(templateSubFolder, header.ImageFile);
            RelatedTemplateFile = Path.Combine(templateSubFolder, RelatedTemplate);
            Level = header.Level;

            // Set the mapping fields
            FieldConfigurations = new List<IConfigurationFieldItem>();

            // Read the configured state transitions if some are defined
            /*
            if (header.StateTransitions != null)
            {
                _stateTransitionField = header.StateTransitions.FieldName;
                _stateTransitions = new Dictionary<string, string>();

                foreach (var transition in header.StateTransitions.Items)
                {
                    _stateTransitions.Add(transition.From, transition.To);
                }
            }
            */

            foreach (var field in header.Fields)
            {
                ConfigurationFieldItemDefaultValue cfidv = null;
                if (field.DefaultValue != null)
                    cfidv = new ConfigurationFieldItemDefaultValue(field.DefaultValue.ShowName,
                                                                   field.DefaultValue.DefaultValue);
                FieldConfigurations.Add(new ConfigurationFieldItem(field.Name, field.MappingName, field.FieldValueType,
                                                                   field.Direction, field.TableRow, field.TableCol,
                                                                   field.TestCaseStepDelimiter, field.HandleAsDocument,
                                                                   field.HandleAsDocumentMode, field.OLEMarkerField,
                                                                   field.OLEMarkerValue, cfidv,
                                                                   field.ShapeOnlyWorkaroundMode,
                                                                   field.TableCol != 0 && field.TableRow != 0,
                                                                   field.WordBookmark,
                                                                   field.VariableName,
                                                                   field.DateTimeFormat
                                                                   ));
            }

            // If the state transition field is not mapped, add it as unmapped field ( so it gets queried anyways)

            /*
            var stateTransition = _fieldConfiguration.FirstOrDefault(x => x.ReferenceFieldName.Equals(_stateTransitionField));
            if (stateTransition == null && _stateTransitionField != null)
            {
                _fieldConfiguration.Add(new ConfigurationFieldItem(_stateTransitionField, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 0, 0, string.Empty, false, HandleAsDocumentType.OleOnDemand, null, false));
            }
             * */

            // Set the converters
            _converters = new Dictionary<string, IConverter>();

            if (header.Converters != null)
            {
                foreach (MappingConverter converter in header.Converters)
                {
                    var valueMapper = new SimpleValueMapper(converter.FieldName);
                    foreach (MappingConverterValue value in converter.Values) valueMapper.AddMapping(value.Text, value.MappingText);
                    _converters.Add(valueMapper.FieldName.ToUpperInvariant(), valueMapper);
                }
            }

            Links = new List<IConfigurationLinkItem>();
            FieldToLinkedItemConfiguration = new List<IConfigurationFieldToLinkedItem>();
            /*
            if (header.Links != null)
            {
                foreach (MappingLink link in header.Links)
                {
                    var cli = new ConfigurationLinkItem(link.LinkValueType, link.Direction, link.TableRow, link.TableCol, link.Overwrite, link.LinkSeparator, link.LinkFormat);
                    _links.Add(cli);
                }
            }

            // Read the 'field to linked work item' configuration.
            if (header.FieldsToLinkedItems != null)
            {
                foreach (var fieldToLinkedItem in header.FieldsToLinkedItems)
                {
                    _fieldToLinkedItemConfiguration.Add(new ConfigurationFieldToLinkedItem(fieldToLinkedItem));
                }
            }
             */
        }

        #endregion Constructors

        #region Private methods

        /// <summary>
        /// Gets an image from a <paramref name="imageFile"/> or returns null if no image was found.
        /// </summary>
        /// <param name="imageFile">File name.</param>
        /// <returns><see cref="Image"/> or null.</returns>
        private Image GetImageFile(string imageFile)
        {
            if (!string.IsNullOrEmpty(imageFile))
            {
                var file = Path.Combine(_configuration.MappingFolder, imageFile);

                if (!File.Exists(file))
                {
                    file = Path.Combine(_configuration.MappingFolder, "../Resources/standard.png");
                }
                if (File.Exists(file))
                {
                    try
                    {
                        using (var fileStream = new FileStream(file, FileMode.Open, FileAccess.Read))
                        {
                            return Image.FromStream(fileStream);
                        }
                    }
                    catch (FileNotFoundException) { }
                    catch (ArgumentException) { }
                }
            }

            return null;
        }

        #endregion Private methods

        #region Implementatn of IConfigurationItem

        /// <summary>
        /// Gets the list of pre-operations.
        /// </summary>
        public IList<IConfigurationTestOperation> PreOperations { get; private set; }

        /// <summary>
        /// Gets the list of post-operations.
        /// </summary>
        public IList<IConfigurationTestOperation> PostOperations { get; private set; }

        /// <summary>
        /// The property gets the information whether the configured item supports the linked items.
        /// </summary>
        public bool ConfigurationSupportsLinkedItem
        {
            get
            {
                return FieldToLinkedItemConfiguration == null || FieldToLinkedItemConfiguration.Count != 0;
            }
        }

        /// <summary>
        /// The work item type name for the destination of the sync process.
        /// </summary>
        public string WorkItemType { get; private set; }

        /// <summary>
        /// Defines reference field name, which must be evaluated to become the sub type of work item.
        /// For example, 'Scenario' is sub type of 'Requirement'.
        /// </summary>
        public string WorkItemSubtypeField { get; private set; }

        /// <summary>
        /// Defines value of related sub type of work item.
        /// For example, 'Scenario' is sub type of 'Requirement'.
        /// </summary>
        public string WorkItemSubtypeValue { get; private set; }

        /// <summary>
        /// The work item type name for the source of the sync process.
        /// </summary>
        public string WorkItemTypeMapping { get; private set; }

        /// <summary>
        /// The regular expression to identify requirement tables with (comparison with cell 1,1).
        /// </summary>
        public string ReqTableIdentifierExpression { get; private set; }

        /// <summary>
        /// The row of the cell to identify requirement tables. See <see cref="ReqTableIdentifierExpression"/> and <see cref="ReqTableCellCol"/>.
        /// </summary>
        public int ReqTableCellRow { get; private set; }

        /// <summary>
        /// The column of the cell to identify requirement tables. See <see cref="ReqTableIdentifierExpression"/> and <see cref="ReqTableCellRow"/>.
        /// </summary>
        public int ReqTableCellCol { get; private set; }

        /// <summary>
        /// The specific <see cref="IConfigurationFieldItem"/> objects which can be accessed by the destination name.
        /// </summary>
        public IList<IConfigurationFieldItem> FieldConfigurations { get; private set; }

        /// <summary>
        /// The property gets all configured fields of type <see cref="IConfigurationFieldToLinkedItem"/> to define linked work items.
        /// </summary>
        public IList<IConfigurationFieldToLinkedItem> FieldToLinkedItemConfiguration { get; private set; }

        /// <summary>
        /// Gets all configured links.
        /// </summary>
        /// <value>All configured links.</value>
        public IList<IConfigurationLinkItem> Links { get; private set; }

        /// <summary>
        /// Image that represents the work item in the UI
        /// </summary>
        public Image ImageFile
        {
            get { return _image ?? (_image = GetImageFile(_imageFilename)); }
        }

        /// <summary>
        /// Related schema of the <see cref="IConfigurationItem"/>.
        /// </summary>
        public string RelatedSchema { get; set; }

        /// <summary>
        /// Related template of the <see cref="IConfigurationItem"/>.
        /// </summary>
        public string RelatedTemplate { get; private set; }

        /// <summary>
        /// Gets template file.
        /// </summary>
        public string RelatedTemplateFile { get; private set; }

        public bool HideElementInWord { get; private set; }

        /// <summary>
        /// This methods can be used to obtain a <see cref="IConverter"/> for a specific <see cref="IField"/>.
        /// The field is identified by the field reference name
        /// The methods uses a dictionary to retrieve the field. Therefore the Access speed is O(1).
        /// </summary>
        /// <param name="fieldName">The <see cref="IField"/> which has to be converted.</param>
        /// <returns>Returns the converter for a specific <see cref="IField"/> object.</returns>
        public IConverter GetConverter(string fieldName)
        {
            if (fieldName == null)
                throw new ArgumentNullException("fieldName");

            if (_converters != null && _converters.ContainsKey(fieldName.ToUpperInvariant()))
                return _converters[fieldName.ToUpperInvariant()];

            return null;
        }

        /// <summary>
        /// Sets the converter for a given field. This can be used to create context sensitive converters
        /// </summary>
        /// <param name="fieldName">The target field of the conversion</param>
        /// <param name="converter">The converter to be used</param>
        public void SetConverter(string fieldName, IConverter converter)
        {
            if (fieldName == null) return;
            if (_converters.ContainsKey(fieldName.ToUpperInvariant()))
                _converters[fieldName.ToUpperInvariant()] = converter;
            else
                _converters.Add(fieldName.ToUpperInvariant(), converter);
        }

        /// <summary>
        /// Checks if configured transitions should be applied: If word item has no manually changed
        /// state field, it will use transitions rules in which case either word item (if mapped) or
        /// TFS item is changed.
        /// </summary>
        /// <param name="workItemTfs">mapped work item in TFS</param>
        /// <param name="workItemWord">mapped work item in word</param>
        public void DoTransition(IWorkItem workItemTfs, IWorkItem workItemWord)
        {
            if (workItemTfs == null || workItemWord == null)
                throw new ArgumentException("An argument was null");

            // dont do anything if nothing is configured
            if (_stateTransitionField == null || _stateTransitions == null) return;

            if (workItemTfs.Fields.Contains(_stateTransitionField))
            {
                var field = workItemTfs.Fields[_stateTransitionField];
                string newState;
                if (!_stateTransitions.TryGetValue(field.Value, out newState)) return;

                if (workItemWord.Fields.Contains(_stateTransitionField))
                {
                    // manual changes overrides automatic transitions!
                    if (!workItemWord.Fields[_stateTransitionField].Value.Equals(field.Value)) return;

                    // Set value to word item because that way it just works natural during later sync
                    SyncServiceTrace.I("WI{0}W:State transition from {1} to {2}", workItemWord.Id, workItemWord.Fields[_stateTransitionField].Value, newState);
                    workItemWord.Fields[_stateTransitionField].Value = newState;
                }
                else
                {
                    // if field is not mapped to word, set its value directly in the TFS item.
                    // It will not get overwritten (because its not mapped...)
                    field.Value = newState;
                }
            }
        }

        /// <summary>
        /// Level of Header configuration
        /// </summary>
        public int Level { get; private set; }

        /// <summary>
        /// Creates a flat clone of the configuration. You can alter the field configuration
        /// of the clone without changing the original configuration.
        /// </summary>
        /// <returns>A flat copy of the configuration</returns>
        public IConfigurationItem Clone()
        {
            return new ConfigurationItem(this);
        }

        #endregion Implementatn of IConfigurationItem
    }
}