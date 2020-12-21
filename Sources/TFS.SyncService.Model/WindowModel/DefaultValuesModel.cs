namespace AIT.TFS.SyncService.Model.WindowModel
{
    #region Usings
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using AIT.TFS.SyncService.Contracts.Configuration;
    using AIT.TFS.SyncService.Contracts.Model;
    using AIT.TFS.SyncService.Factory;
    using AIT.TFS.SyncService.Model.WindowModelBase;
    #endregion

    /// <summary>
    /// This class implements a model for default value control to set the default values and bind list view column widths.
    /// </summary>
    public class DefaultValuesModel : ExtBaseModel
    {
        #region Private fields
        
        private readonly DefaultValueModelCollection _defValsList = new DefaultValueModelCollection();
        private bool _isDirty;
        private double _column1Width;
        private double _column2Width;

        #endregion Private fields
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultValuesModel"/> class.
        /// </summary>
        /// <param name="model">Associated model of document to show the default values from.</param>
        public DefaultValuesModel(ISyncServiceDocumentModel model)
        {
            if (model == null)
                throw new ArgumentNullException("model");
            DocumentModel = model;
            RefreshDefaultValues();
        }

        #endregion Constructors
        #region Public properties

        /// <summary>
        /// Gets a model of associated word document.
        /// </summary>
        public ISyncServiceDocumentModel DocumentModel { get; private set; }

        /// <summary>
        /// Gets or sets the flag if the at least one of all default values changed (stored in active document).
        /// </summary>
        public bool AtLeastOneValueChanged
        {
            get { return _isDirty; }
            set
            {
                if (_isDirty != value)
                {
                    _isDirty = value;
                    TriggerPropertyChanged(this, nameof(AtLeastOneValueChanged));
                }
            }
        }

        /// <summary>
        /// Gets or sets the value for the first column width.
        /// </summary>
        public double Column1Width
        {
            get { return _column1Width; }

            set
            {
                if (Math.Abs(this._column1Width - value) > double.Epsilon)
                {
                    _column1Width = value;
                    TriggerPropertyChanged(this, nameof(Column1Width));
                }
            }
        }

        /// <summary>
        /// Gets or sets the value for the second column width.
        /// </summary>
        public double Column2Width
        {
            get { return _column2Width; }

            set
            {
                if (Math.Abs(this._column2Width - value) > double.Epsilon)
                {
                    _column2Width = value;
                    TriggerPropertyChanged(this, nameof(Column2Width));
                }
            }
        }

        /// <summary>
        /// Gets the list of all defined default values from mapping bundle.
        /// </summary>
        public DefaultValueModelCollection DefaultValues
        {
            get { return _defValsList; }
        }

        #endregion Public properties
        #region Public methods

        /// <summary>
        /// Method will reread the default values from active document.
        /// </summary>
        public void RefreshDefaultValues()
        {
            BuildDefaultValueList(false);
        }

        /// <summary>
        /// Method resets all default values in active document. It means, that all default values removed from active document.
        /// </summary>
        public void ResetDefaultValues()
        {
            BuildDefaultValueList(true);
        }

        #endregion Public methods
        #region Private methods
        /// <summary>
        /// Method builds the default value list.
        /// </summary>
        /// <param name="resetValuesInDocument">If true, all default values from document removed.</param>
        private void BuildDefaultValueList(bool resetValuesInDocument)
        {
            AtLeastOneValueChanged = false;
            _defValsList.Clear();

            var service = SyncServiceFactory.GetService<IConfigurationService>();
            if (service == null)
                return;

            var ignoreFormatting = DocumentModel.Configuration.IgnoreFormatting;
            var conflictOverwrite = DocumentModel.Configuration.ConflictOverwrite;
            service.GetConfiguration(DocumentModel.WordDocument).ActivateMapping(DocumentModel.MappingShowName);
            DocumentModel.Configuration.IgnoreFormatting = ignoreFormatting;
            DocumentModel.Configuration.ConflictOverwrite = conflictOverwrite;

            foreach (IConfigurationItem configurationItem in service.GetConfiguration(DocumentModel.WordDocument).GetConfigurationItems())
            {
                int defaultValueCounter = 0;

                foreach (IConfigurationFieldItem configurationFieldItem in configurationItem.FieldConfigurations)
                {
                    if (configurationFieldItem.DefaultValue == null) continue;

                    if (!string.IsNullOrEmpty(configurationFieldItem.DefaultValue.DefaultValue))
                    {
                        defaultValueCounter++;
                    }
                }

                if (defaultValueCounter <= 0) continue;

                var title
                    = new DefaultValueModel(string.Empty, configurationItem.WorkItemTypeMapping, string.Empty, "Hidden",
                                            configurationItem.WorkItemType);

                _defValsList.Add(title);

                IList<IConfigurationFieldItem> fields = configurationItem.FieldConfigurations;
                if (fields == null || fields.Count == 0)
                    return;
                foreach (IConfigurationFieldItem field in fields)
                {
                    if (field.DefaultValue == null)
                        continue;
                    string defaultValue = field.DefaultValue.DefaultValue;
                    if (DocumentModel != null)
                    {
                        if (resetValuesInDocument)
                        {
                            DocumentModel.ClearDefaultValue(field.SourceMappingName);
                        }
                        else if (DocumentModel.GetFieldDefaultValue(field.ReferenceFieldName, configurationItem.WorkItemTypeMapping) != null)
                        {

                            defaultValue = DocumentModel.GetFieldDefaultValue(field.ReferenceFieldName, configurationItem.WorkItemTypeMapping);
                            AtLeastOneValueChanged = true;
                        }
                    }
                    if (defaultValue == null)
                        defaultValue = string.Empty;
                    var oneDefaultValue
                        = new DefaultValueModel(field.ReferenceFieldName, field.DefaultValue.ShowName, defaultValue,
                                                "Visible", configurationItem.WorkItemTypeMapping);
                    oneDefaultValue.PropertyChanged += HandleDefaultValueModelPropertyChanged;
                    _defValsList.Add(oneDefaultValue);
                    if (DocumentModel != null)
                        DocumentModel.SetFieldDefaultValue(
                            oneDefaultValue.ReferenceFieldName, oneDefaultValue.Value, oneDefaultValue.WorkItemType);
                }
            }
        }

        #endregion Private methods
        #region Private event handler methods

        /// <summary>
        /// Called when a property value changes.
        /// </summary>
        /// <param name="sender">Sender of the change.</param>
        /// <param name="e">A <see cref="PropertyChangedEventArgs"/> contains the event data.</param>
        private void HandleDefaultValueModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            AtLeastOneValueChanged = true;
            var defaultValueModel = sender as DefaultValueModel;
            if (defaultValueModel != null)
                DocumentModel.SetFieldDefaultValue(defaultValueModel.ReferenceFieldName, defaultValueModel.Value, defaultValueModel.WorkItemType);
        }

        #endregion Private event handler methods
    }
}