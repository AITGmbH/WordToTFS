#region Usings
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Adapter;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Properties;
using AIT.TFS.SyncService.Service.InfoStorage;
using AIT.TFS.SyncService.Service.Utils;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
#endregion

namespace AIT.TFS.SyncService.Model
{
    /// <summary>
    /// The class implements <see cref="ISyncServiceDocumentModel"/>.
    /// </summary>
    public class SyncServiceDocumentModel : ISyncServiceDocumentModel
    {
        #region Private consts

        private const string ConstFieldPrefix = "__prefix__";
        private const string ConstTfsBound = "TFSBound";
        private const string ConstTfsProject = "TFSProject";
        private const string ConstTfsServer = "TFSServer";
        private const string ConstMappingShowName = "MappingShowName";
        private const string ConstQcQueryPath = "QCQueryPath";
        private const string ConstQcUseLinkedWorkItems = "QCUseLinkedWorkItems";
        private const string ConstQcLinkTypes = "QCLinkTypes";
        private const string ConstAreaIterationPath = "AreaIterationPath";
        private const string ConstTestReportGenerated = "WordToTFS/TestReport/ReportGenerated";
        private const string ConstTestReportRunning = "WordToTFS/TestReport/ReportRunning";
        private const string ConstTestReportData = "WordToTFS/TestReport/{0}";

        #endregion Private consts

        #region Private fields

        private bool _operationInProgress;

        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncServiceDocumentModel"/> class.
        /// </summary>
        /// <param name="wordDocument">Associated word document.</param>
        public SyncServiceDocumentModel(Document wordDocument)
        {
            Guard.ThrowOnArgumentNull(wordDocument, "wordDocument");

            WordDocument = wordDocument;
            Load();
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the associated document.
        /// </summary>
        private Document WordDocument { get; set; }

        #endregion Public properties

        #region Private methods

        /// <summary>
        /// Publish the information that the connection settings are wrong
        /// </summary>
        /// <param name="text"></param>
        /// <param name="explanation"></param>
        private static void PublishConnectionError(string text, string explanation)
        {
            var info = new UserInformation
                       {
                Text = text,
                Explanation = string.Format(explanation),
                Type = UserInformationType.Error,

            };
            SyncServiceTrace.W("{0}:{1}", info.Text, info.Explanation);

            var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
            if (infoStorage != null)
                infoStorage.AddItem(info);

        }

        /// <summary>
        /// Loads data from the document.
        /// </summary>
        private void Load()
        {
            MappingShowName = GetDocumentVariableValue(ConstMappingShowName, null);
        }

        /// <summary>
        /// Returns a variable saved in the word file
        /// </summary>
        /// <param name="variableName">Variable name</param>
        /// <param name="workItemType">Work item type</param>
        /// <returns>Variable value</returns>
        private string GetDocumentVariableValue(string variableName, string workItemType)
        {
            for (var index = 0; index < 6; index++)
            {
                try
                {
                    foreach (Variable var in WordDocument.Variables)
                    {
                        if (var == null)
                            continue;

                        var combinedValues = var.Name.Split(';');

                        if (combinedValues.Length == 2)
                        {
                            var workItemTypeValue = combinedValues[0];
                            var nameValue = combinedValues[1];

                            if (string.IsNullOrEmpty(workItemTypeValue))
                            {
                                if (var.Name == variableName && var.Value != null)
                                {
                                    return var.Value;
                                }
                            }
                            else
                            {
                                if (workItemTypeValue.Equals(workItemType) && nameValue.Equals(variableName) && var.Value != null)
                                {
                                   return var.Value;
                                }
                            }
                        }
                        else if (combinedValues.Length == 1)
                        {
                            if (var.Name == variableName && var.Value != null)
                            {
                                    return var.Value;
                            }
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
                catch (COMException ex)
                {
                    // MIK: ???
                    // ApplicationBusy or something like that? Trying to figure out what is catched here:
                    // If you never see the next assert hit by July 2014 you can probably remove try/catch -SER
                    Debug.Assert(ex != null, "Check sourcecode comment.");
                    Thread.Sleep(500);
                    if (index == 5) throw;
                }
            }
            return null;
        }

        /// <summary>
        /// Method removes the variable from the active document.
        /// </summary>
        /// <param name="variableName">Variable to remove if exists.</param>
        private void ClearDocumentVariableValue(string variableName)
        {
            foreach (Variable variable in WordDocument.Variables)
            {
                if (variable == null)
                    continue;
                if (variable.Name != variableName)
                    continue;
                // Set to null - variable will be removed
                variable.Value = null;
                OnDocumentVariableChanged();
                break;
            }
        }

        /// <summary>
        /// Method sets the variable value in active document.
        /// </summary>
        /// <param name="variableName">Variable to set the value for.</param>
        /// <param name="variableValue">Value to set in the variable.</param>
        /// <param name="workItemType">Work item type of the variable.</param>
        private void SetDocumentVariableValue(string variableName, string variableValue, string workItemType)
        {
            if (!string.IsNullOrEmpty(workItemType))
                workItemType += ";";

            var combinedName = workItemType + variableName;
            ClearDocumentVariableValue(combinedName);
            object value = variableValue;
            OnDocumentVariableChanged();
            WordDocument.Variables.Add(combinedName, ref value);
        }

        #endregion Private methods

        #region Private raise event methods

        /// <summary>
        /// The method raises event <see cref="TestReportDataChanged"/>.
        /// </summary>
        private void OnTestReportDataChanged()
        {
            TestReportDataChanged?.Invoke(this, new SyncServiceEventArgs(WordDocument.Name, WordDocument.FullName));
        }

        /// <summary>
        /// Method raises the <see cref="DocumentVariableChanged"/> event.
        /// </summary>
        private void OnDocumentVariableChanged()
        {
            DocumentVariableChanged?.Invoke(this, new SyncServiceEventArgs(WordDocument.Name, WordDocument.FullName));
        }

        private void OnPropertyChangedEvent(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion Private raise event methods

        #region Implementation of ISyncServiceDocumentModel

        #region Public Properties
        /// <summary>
        /// Gets or sets a value indicating whether an operation is running.
        /// </summary>
        public bool OperationInProgress
        {
            get
            {
                return _operationInProgress;
            }
            set
            {
                if (_operationInProgress != value)
                {
                    _operationInProgress = value;
                    OnPropertyChangedEvent(nameof(OperationInProgress));
                }
            }
        }

        /// <summary>
        /// Gets the associated document.
        /// </summary>
        object ISyncServiceDocumentModel.WordDocument
        {
            get
            {
                return WordDocument;
            }
        }

        /// <summary>
        /// Gets the name of associated document. If associated document is not valid string.Empty returned.
        /// </summary>
        public string DocumentName
        {
            get
            {
                return WordDocument.Name;
            }
        }

        /// <summary>
        /// Gets the flag that tells the active document is bound to TFS project.
        /// </summary>
        public bool TfsDocumentBound
        {
            get
            {
                return GetDocumentVariableValue(ConstTfsBound, null) != null;
            }
        }

        /// <summary>
        /// Gets the TFS server where active document is bound. See the <see cref="TfsDocumentBound"/>.
        /// </summary>
        public string TfsServer
        {
            get { return GetDocumentVariableValue(ConstTfsServer, null); }
        }

        /// <summary>
        /// Gets the TFS project where active document is bound. See the <see cref="TfsDocumentBound"/>.
        /// </summary>
        public string TfsProject
        {
            get { return GetDocumentVariableValue(ConstTfsProject, null); }
        }

        /// <summary>
        /// Gets the show name of mapping to use with active document.
        /// </summary>
        /// <remarks>
        /// The mapping show name encapsulates many information. For example mapping file, template file.
        /// </remarks>
        public string MappingShowName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets into document Area or Iteration path.
        /// </summary>
        public string AreaIterationPath
        {
            get
            {
                return GetDocumentVariableValue(ConstAreaIterationPath, null);
            }
            set
            {
                SetDocumentVariableValue(ConstAreaIterationPath, value, null);
            }
        }

        /// <summary>
        /// Gets the configuration for the currently active document
        /// </summary>
        public IConfiguration Configuration
        {
            get
            {
                var configurationService = SyncServiceFactory.GetService<IConfigurationService>();
                if (configurationService == null)
                {
                    Debug.Assert(configurationService == null, "Configuration service not exists!");
                    return null;
                }
                return configurationService.GetConfiguration(WordDocument);
            }
        }
        #endregion
        #region Public methods

        /// <summary>
        /// Read query configuration properties from document.
        /// </summary>
        /// <param name="queryConfiguration">IQueryConfiguration instance.</param>
        public void ReadQueryConfiguration(IQueryConfiguration queryConfiguration)
        {
            if (queryConfiguration == null)
                return;

            // get query path
            var queryPath = GetDocumentVariableValue(ConstQcQueryPath, null);
            if (!string.IsNullOrEmpty(queryPath))
            {
                queryConfiguration.QueryPath = queryPath;
            }

            // get UseLinkedWorkItems flag
            var useLinkedWorkItems = GetDocumentVariableValue(ConstQcUseLinkedWorkItems, null);
            if (!string.IsNullOrEmpty(useLinkedWorkItems))
            {
                queryConfiguration.UseLinkedWorkItems = bool.Parse(useLinkedWorkItems);
            }

            // get link types
            var linkTypes = GetDocumentVariableValue(ConstQcLinkTypes, null);
            if (!string.IsNullOrEmpty(linkTypes))
            {
                queryConfiguration.LinkTypes.Clear();
                var links = linkTypes.Split(';');
                foreach (var link in links)
                {
                    queryConfiguration.LinkTypes.Add(link);
                }
            }
        }

        /// <summary>
        /// Saves data to document
        /// </summary>
        public void Save()
        {
            SetDocumentVariableValue(ConstMappingShowName, MappingShowName, null);
        }

        /// <summary>
        /// Save query configuration properties to document.
        /// </summary>
        /// <param name="queryConfiguration">IQueryConfiguration instance.</param>
        public void SaveQueryConfiguration(IQueryConfiguration queryConfiguration)
        {
            if (queryConfiguration == null)
                return;

            // save query path
            SetDocumentVariableValue(ConstQcQueryPath, queryConfiguration.QueryPath, null);

            // save UseLinkedWorkItems flag
            SetDocumentVariableValue(ConstQcUseLinkedWorkItems, queryConfiguration.UseLinkedWorkItems.ToString(), null);

            // save Link types
            var linkTypes = string.Empty;
            foreach (var linkType in queryConfiguration.LinkTypes)
                linkTypes += linkType + ';';
            SetDocumentVariableValue(ConstQcLinkTypes, linkTypes, null);
        }

        /// <summary>
        /// Clears the default value for given field.
        /// </summary>
        /// <param name="fieldName">Name of the field to clear the default value for.</param>
        public void ClearDefaultValue(string fieldName)
        {
            // We use a special prefix to be sure that no conflict occurs with 'our' data.
            ClearDocumentVariableValue(ConstFieldPrefix + fieldName);
        }

        /// <summary>
        /// Methods sets the desired default value for given field in active document.
        /// </summary>
        /// <param name="fieldName">Name of the field to store the default value for.</param>
        /// <param name="fieldDefaultValue">Default value to store for the given field.</param>
        /// <param name="workItemType">Work item type of the value.</param>
        public void SetFieldDefaultValue(string fieldName, string fieldDefaultValue, string workItemType)
        {
            SetDocumentVariableValue(ConstFieldPrefix + fieldName, fieldDefaultValue, workItemType);
        }

        /// <summary>
        /// Method gets the stored default value for given field.
        /// </summary>
        /// <param name="fieldName">Name of the field to get the default value for.</param>
        /// <param name="workItemType">Work item type of the value.</param>
        /// <returns>Default value. Null if not defined.</returns>
        public string GetFieldDefaultValue(string fieldName, string workItemType)
        {
            return GetDocumentVariableValue(ConstFieldPrefix + fieldName, workItemType);
        }

        /// <summary>
        /// Call this method to bind the active document to the TFS project.
        /// </summary>
        public void BindProject()
        {
            if (TfsDocumentBound)
                return;

            var tfs = SyncServiceFactory.GetService<ITeamFoundationServerMaintenance>();
            if (tfs == null)
            {
                return;
            }

            var serverUrl = string.Empty;
            var projectName = string.Empty;

            if (!string.IsNullOrEmpty(Configuration.DefaultProjectName) && !string.IsNullOrEmpty(Configuration.DefaultServerUrl))
            {
                if (!tfs.ValidateConnectionSettings(Configuration.DefaultServerUrl, Configuration.DefaultProjectName, out serverUrl, out projectName))
                {
                    //Publish error if connection fails
                    PublishConnectionError(Resources.Error_Wrong_Connection_Short, Resources.Error_Wrong_Connection_Long);
                    return;
                }
            }
            else
            {
                try
                {
                    if (!tfs.SelectOneTfsProject(out serverUrl, out projectName))
                    {
                        return;
                    }
                }
                catch (Exception e)
                {
                    PublishConnectionError(Resources.Error_TeamPickerFailed_Short, Resources.Error_TeamPickerFailed_Long);
                    SyncServiceTrace.W(Resources.Error_TeamPickerFailed_Long);
                    SyncServiceTrace.W(e.Message);
                    return;
                }
            }

            SetDocumentVariableValue(ConstTfsBound, ConstTfsBound, null);
            SetDocumentVariableValue(ConstTfsServer, serverUrl, null);
            SetDocumentVariableValue(ConstTfsProject, projectName, null);
        }

        /// <summary>
        /// Bind the project to the given server and project
        /// </summary>
        /// <param name="tfsServer"></param>
        /// <param name="tfsProject"></param>
        public void BindProject(string tfsServer, string tfsProject)
        {
            if (TfsDocumentBound)
                return;

            SetDocumentVariableValue(ConstTfsBound, ConstTfsBound, null);
            SetDocumentVariableValue(ConstTfsServer, tfsServer, null);
            SetDocumentVariableValue(ConstTfsProject, tfsProject, null);
        }

        /// <summary>
        /// Call this method to unbound the active document from the TFS project.
        /// </summary>
        public void UnbindProject()
        {
            if (!TfsDocumentBound)
                return;

            ClearDocumentVariableValue(ConstTfsBound);
            ClearDocumentVariableValue(ConstTfsServer);
            ClearDocumentVariableValue(ConstTfsProject);
        }

        /// <summary>
        /// Verifies this instance.
        /// </summary>
        public IList<IConfigurationFieldItem> GetMissingFields()
        {
            var serverUrl = GetDocumentVariableValue(ConstTfsServer, null);
            var projectName = GetDocumentVariableValue(ConstTfsProject, null);
            if (string.IsNullOrEmpty(serverUrl) || string.IsNullOrEmpty(projectName))
            {
                return new List<IConfigurationFieldItem>();
            }

            // verify template
            var veryfier = new TemplateVerifier();
            var configuration = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(WordDocument);
            var adapter = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(serverUrl, projectName, null, configuration);

            if (!veryfier.VerifyTemplateMapping(this, (ITfsService)adapter, configuration))
            {
                return veryfier.MissingFields;
            }
            return new List<IConfigurationFieldItem>();
        }

        /// <summary>
        /// Verifies this instance.
        /// </summary>
        public IList<IConfigurationFieldItem> GetWronglyMappedFields()
        {
            var serverUrl = GetDocumentVariableValue(ConstTfsServer, null);
            var projectName = GetDocumentVariableValue(ConstTfsProject, null);
            if (string.IsNullOrEmpty(serverUrl) || string.IsNullOrEmpty(projectName))
            {
                return new List<IConfigurationFieldItem>();
            }

            // verify template
            var veryfier = new TemplateVerifier();
            var configuration = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(WordDocument);
            var adapter = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(serverUrl, projectName, null, configuration);

            if (!veryfier.VerifyTemplateMapping(this, (ITfsService)adapter, configuration))
            {
                return veryfier.WrongMappedFields;
            }
            return new List<IConfigurationFieldItem>();
        }

        #endregion

        /// <summary>
        /// Occurs when variable changed in any document.
        /// </summary>
        public event EventHandler<SyncServiceEventArgs> DocumentVariableChanged;

        #region Test report related

        /// <summary>
        /// Occurs if any data in test report area was changed.
        /// </summary>
        public event EventHandler<SyncServiceEventArgs> TestReportDataChanged;

        /// <summary>
        /// Gets or sets the flag telling if the test report is running.
        /// </summary>
        public bool TestReportRunning
        {
            get
            {
                // null or Empty is false. Any text is true.
                return !string.IsNullOrEmpty(GetDocumentVariableValue(ConstTestReportRunning, null));
            }
            set
            {
                OperationInProgress = value;
                var varValue = value ? "true" : string.Empty;
                SetDocumentVariableValue(ConstTestReportRunning, varValue, null);
                OnTestReportDataChanged();
            }
        }

        /// <summary>
        /// Gets or sets the flag telling if the document contains the test report - specification report or result report.
        /// </summary>
        /// <remarks>Value is stored in document variable.</remarks>
        public bool TestReportGenerated
        {
            get
            {
                // null or Empty is false. Any text is true.
                return !string.IsNullOrEmpty(GetDocumentVariableValue(ConstTestReportGenerated, null));
            }
            set
            {
                var varValue = value ? "true" : string.Empty;
                SetDocumentVariableValue(ConstTestReportGenerated, varValue, null);
                OnTestReportDataChanged();
            }
        }

        /// <summary>
        /// The method reads the data associated with given key from document.
        /// </summary>
        /// <param name="key">Key to get the data for.</param>
        /// <returns>Read data. <c>null</c> if data was not written before.</returns>
        public string ReadTestReportData(string key)
        {
            return GetDocumentVariableValue(string.Format(CultureInfo.InvariantCulture, ConstTestReportData, key), null);
        }

        /// <summary>
        /// The method writes the given data with given associated key to document.
        /// </summary>
        /// <param name="key">Key to store the data for.</param>
        /// <param name="data">Data to store.</param>
        public void WriteTestReportData(string key, string data)
        {
            SetDocumentVariableValue(string.Format(CultureInfo.InvariantCulture, ConstTestReportData, key), data, null);
        }

        /// <summary>
        /// The method clears the data associated with given key from document.
        /// </summary>
        /// <param name="key">Key to clear the associated data.</param>
        public void ClearTestReportData(string key)
        {
            ClearDocumentVariableValue(string.Format(CultureInfo.InvariantCulture, ConstTestReportData, key));
        }

        #endregion Test report related

        #endregion Implementation of ISyncServiceDocumentModel

        #region Implementation of INotifyPropertyChanged

        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Implementation of INotifyPropertyChanged
    }
}
