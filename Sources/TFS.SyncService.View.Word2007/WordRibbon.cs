using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Deployment.Application;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Threading;
using AIT.TFS.SyncService.Adapter.TFS2012;
using AIT.TFS.SyncService.Adapter.Word2007;
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
using AIT.TFS.SyncService.Common.Helper;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Contracts.TemplateManager;
using AIT.TFS.SyncService.Contracts.Word;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Model.TemplateManagement;
using AIT.TFS.SyncService.Model.TestReport;
using AIT.TFS.SyncService.Model.WindowModel;
using AIT.TFS.SyncService.Service.Configuration;
using AIT.TFS.SyncService.Service.Utils;
using AIT.TFS.SyncService.View.Controls;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using AIT.TFS.SyncService.View.Controls.TemplateManager;
using AIT.TFS.SyncService.View.Controls.TestReport;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using TFS.SyncService.View.Word2007.Controls;
using TFS.SyncService.View.Word2007.Properties;
using IWin32Window = System.Windows.Forms.IWin32Window;
using MessageBox = System.Windows.Forms.MessageBox;
using MessageBoxOptions = System.Windows.Forms.MessageBoxOptions;
using AIT.TFS.SyncService.Contracts.Adapter;

namespace TFS.SyncService.View.Word2007
{
    /// <summary>
    /// 'AIT.WordToTFS' ribbon for word - shows all required UI elements.
    /// </summary>
    public partial class WordRibbon : RibbonBase, IWordRibbon, IWin32Window
    {
        #region Private structures

        /// <summary>
        /// Used to pass information to the background worker since that can only take one argument.
        /// </summary>
        private struct PublishArguments
        {
            public IEnumerable<IWorkItem> WorkItems;

            public bool ForcePublish;

            public ISyncServiceDocumentModel DocumentModel;
        }

        #endregion Private structures

        #region Private fields

        private readonly TemplateManager _templateManager;
        private readonly Dispatcher _dispatcher;
        private PublishingStateControl _publishingStateControl;
        private const string WordClassName = "OpusApp";
        private IWordSyncAdapter _refreshWordAdapter;
        private IWordSyncAdapter _publishWordAdapter;

        #endregion Fields

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="WordRibbon"/> class.
        /// </summary>
        public WordRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            _dispatcher = Dispatcher.CurrentDispatcher;

            EmbeddedResources.ExportResources();
            TraceLevel level = (TraceLevel)Settings.Default.EnableDebugging;
            SyncServiceTrace.DebugLevel.Level = level;
            SettingsModel.Level = level;

            _templateManager = new TemplateManager();
            _templateManager.TemplatesChanged += (s, a) => _dispatcher.BeginInvoke((Action)(() => UpdateTemplateSelection(WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument))));
        }

        /// <summary>
        /// Register to services after they have been initialized
        /// </summary>
        public void RegisterToServices()
        {
            var infoService = SyncServiceFactory.GetService<IInfoStorageService>();
            infoService.UserInformation.CollectionChanged += PublishingInformationOnCollectionChanged;

            SyncServiceFactory.RegisterService(Dispatcher.CurrentDispatcher);
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            if (progressService != null)
            {
                progressService.ShowHideDialogChanged += (s, e) => ToggleProgressWindowVisibility();
            }
        }

        /// <summary>
        /// Shows the error window if an information is published to the information storage
        /// </summary>
        private void PublishingInformationOnCollectionChanged(object sender, NotifyCollectionChangedEventArgs notifyCollectionChangedEventArgs)
        {
            var collection = sender as IInfoCollection<IUserInformation>;
            if (collection == null) return;
            if (Globals.ThisAddIn.Application.Documents.Count == 0) return;

            // TODO Information storage should be bound to a model. This currently shows errors from any action in whatever is the currently active document
            var document = Globals.ThisAddIn.Application.ActiveDocument;

            // TODO Set a context specific ErrorView - Title like "The following errors occured during <context>"
            _dispatcher.Invoke(() =>
                {
                    // dont show if publish control is shown
                    if (Globals.ThisAddIn.TaskPaneVisible<PublishingStateControl>(document))
                    {
                        return;
                    }

                    var errorPane = GetHostPaneForDocument<ErrorWindow>(document);
                    var errorView = errorPane == null ? new ErrorWindow(new ErrorModel(Resources.ErrorPanelText, _dispatcher), document) : errorPane.HostedControl;
                    SetHostPaneVisibility(document, collection.Count > 0, errorView, Resources.ErrorPanelCaption);
                });
        }

        #endregion Constructor

        #region Public properties

        /// <summary>
        /// Gets the word application model.
        /// </summary>
        public ISyncServiceModel WordApplicationModel { get; private set; }

        #endregion Public properties

        #region Private methods

        /// <summary>
        /// Load default values into a work item or header table
        /// </summary>
        /// <param name="item">Configuration from which to take default values.</param>
        /// <param name="documentModel">Document model from which to take default values that override default values from configuration.</param>
        /// <param name="workItem">Work item for which to set default values.</param>
        private static void SetDefaultValues(IConfigurationItem item, ISyncServiceDocumentModel documentModel, IWorkItem workItem)
        {
            foreach (var field in item.FieldConfigurations)
            {
                if (field.DefaultValue == null || field.Direction == Direction.SetInNewTfsWorkItem)
                {
                    continue;
                }

                var defaultValue = field.DefaultValue.DefaultValue;

                var newDefaultValue = documentModel.GetFieldDefaultValue(field.ReferenceFieldName, item.WorkItemTypeMapping);
                if (newDefaultValue != null)
                {
                    defaultValue = newDefaultValue;
                }

                if (string.IsNullOrEmpty(defaultValue))
                {
                    continue;
                }

                try
                {
                    workItem.Fields[field.ReferenceFieldName].Value = defaultValue;
                }
                catch (COMException exception)
                {
                    SyncServiceTrace.LogException(exception);
                }
            }
        }

        /// <summary>
        /// Load allowed values for dropdown lists when adding new items to document.
        /// </summary>
        /// <param name="documentModel">DocumentModel with server data.</param>
        /// <param name="workItem">WorkItem or header for to set allowed values.</param>
        /// <param name="loadAllValues">Sets whether to load all allowed values.</param>
        private static void LoadAllowedValues(ISyncServiceDocumentModel documentModel, IWorkItem workItem, bool loadAllValues)
        {
            if (documentModel.TfsDocumentBound == false)
            {
                return;
            }

            foreach (var field in workItem.Fields)
            {
                field.AllowedValues = new List<string>(new[] { "Loading..." });
            }

            var tfsAdapter = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(documentModel.TfsServer, documentModel.TfsProject, null, documentModel.Configuration);

            if (loadAllValues)
            {
                foreach (var field in workItem.Fields)
                {
                    field.AllowedValues = ((ITfsService)tfsAdapter).FieldDefinitions[field.ReferenceName].AllowedValues.Cast<string>().ToList();
                }
            }
            else
            {
                tfsAdapter.Open(null);
                var tfsWorkItem = tfsAdapter.CreateNewWorkItem(workItem.Configuration);

                foreach (var field in workItem.Fields)
                {
                    if(field.Configuration.FieldValueType != FieldValueType.BasedOnVariable)
                        field.AllowedValues = tfsWorkItem.Fields[field.ReferenceName].AllowedValues;
                }
            }
        }

        /// <summary>
        /// The method shows hidden / hides showed progress window.
        /// </summary>
        private void ToggleProgressWindowVisibility()
        {
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            if (progressService == null)
            {
                return;
            }

            if (progressService.IsVisibleProgressWindow)
            {
                ProgressWindow.HideWindow();
            }
            else
            {
                ProgressWindow.ShowWindow(Handle);
            }
        }

        /// <summary>
        /// The method activates template.
        /// </summary>
        /// <param name="documentModel">Related document.</param>
        /// <param name="mappingShowName">The new template to select.</param>
        private void ActivateTemplate(ISyncServiceDocumentModel documentModel, string mappingShowName)
        {
            SetHostPaneVisibility<DefaultValuesControl>(documentModel.WordDocument as Document, false);

            // find the template in the dropdown list and select it.
            var itemToSelect = dropDownSelectTemplate.Items.FirstOrDefault(x => x.Label.Equals(mappingShowName, StringComparison.OrdinalIgnoreCase));
            if (itemToSelect != null)
            {
                if (itemToSelect.Label != documentModel.MappingShowName)
                {
                    SyncServiceTrace.I("Activating template {0}", documentModel.MappingShowName);
                    documentModel.MappingShowName = itemToSelect.Label;
                }
            }
            else
            {
                return;
            }

            dropDownSelectTemplate.SelectedItem = itemToSelect;
            documentModel.Configuration.ActivateMapping(dropDownSelectTemplate.SelectedItem.Label);

            // Fix for 15890. Warn users when selecting misconfigured templates
            if (!documentModel.Configuration.GetConfigurationItems().All(x => x.FieldConfigurations.Any(y => y.ReferenceFieldName.Equals("System.Id", StringComparison.OrdinalIgnoreCase))) ||
                !documentModel.Configuration.GetConfigurationItems().All(x => x.FieldConfigurations.Any(y => y.ReferenceFieldName.Equals("System.Rev", StringComparison.OrdinalIgnoreCase))))
            {
                MessageBox.Show(this, Resources.Error_MissingFields, Resources.Error_MessageBoxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, 0);
            }

            chkConflictOverwrite.Visible = documentModel.Configuration.EnableSwitchConflictOverwrite;
            chkConflictOverwrite.Checked = documentModel.Configuration.ConflictOverwrite;
            chkIgnoreFormatting.Visible = documentModel.Configuration.EnableSwitchFormatting;
            chkIgnoreFormatting.Checked = documentModel.Configuration.IgnoreFormatting;

            LoadWorkItemDropdownMenu(documentModel, menuNewEmptyWorkItem, Resources.WorkItemPrefix_Empty, HandleMenuNewEmptyWorkItemClick);
            LoadWorkItemDropdownMenu(documentModel, menuNewWorkItem, Resources.WorkItemPrefix_New, HandleMenuNewWorkItemClick);

            // Load configured headers if any
            menuNewHeader.Items.Clear();
            if (documentModel.Configuration.Headers != null)
            {
                foreach (var configurationItem in documentModel.Configuration.Headers)
                {
                    var ribbonButton = Factory.CreateRibbonButton();

                    ribbonButton.Tag = configurationItem;
                    ribbonButton.Label = Resources.WorkItemPrefix_New + configurationItem.WorkItemType;
                    ribbonButton.Image = configurationItem.ImageFile;
                    ribbonButton.Click += HandleMenuNewHeaderClick;

                    menuNewHeader.Items.Add(ribbonButton);
                }
            }

            EnableRibbonControls(documentModel);
        }

        /// <summary>
        /// Updates the template selection in case the list of configured templates has changed.
        /// </summary>
        /// <param name="documentModel">The document for which to update template selection.</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        private void UpdateTemplateSelection(ISyncServiceDocumentModel documentModel)
        {
            UpdateTemplateDropDown(documentModel);

            var boundTemplateItem = dropDownSelectTemplate.Items.FirstOrDefault(x => x.Label == documentModel.MappingShowName);
            var defaultTemplateItem = dropDownSelectTemplate.Items.FirstOrDefault(x => x.Label == documentModel.Configuration.DefaultMapping);

            if (boundTemplateItem != null)
            {
                ActivateTemplate(documentModel, boundTemplateItem.Label);
            }
            else if (defaultTemplateItem != null)
            {
                ActivateTemplate(documentModel, defaultTemplateItem.Label);
            }
            else if (dropDownSelectTemplate.Items.Any())
            {
                ActivateTemplate(documentModel, dropDownSelectTemplate.Items.First().Label);
            }
        }

        /// <summary>
        /// Reload all available templates and fill them into the dropdown ComboBox. Different documents may have different templates available for example if they are connected to different projects.
        /// </summary>
        /// <param name="documentModel">The document for which to load available templates.</param>
        private void UpdateTemplateDropDown(ISyncServiceDocumentModel documentModel)
        {
            SyncServiceTrace.D("Creating template dropdown...");
            WordApplicationModel.CleanUpModelList();
            dropDownSelectTemplate.Enabled = false;

            documentModel.Configuration.RefreshMappings();
            documentModel.Configuration.ActivateAllMappings();
            dropDownSelectTemplate.Items.Clear();

            foreach (var template in _templateManager.AvailableTemplates)
            {
                // Add all available. Disabled can be added only if in document selected.
                if (template.TemplateState == TemplateState.Available || (template.TemplateState == TemplateState.Disabled && template.ShowName == documentModel.MappingShowName))
                {
                    // Check if the template has been marked as beeing only available when connect to a project of a certain name.
                    var uriComponents = documentModel.TfsServer == null ? new[] { string.Empty } : documentModel.TfsServer.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                    var isProjectMatched = template.ProjectName == null || template.ProjectName.Equals(documentModel.TfsProject);
                    var isProjectCollectionMatched = template.ProjectCollectionName == null || (uriComponents.Length == 2 && template.ProjectCollectionName.Equals(uriComponents[1]));
                    var isServerMatched = template.ServerName == null || (uriComponents.Length == 2 && template.ServerName.Equals(uriComponents[0]));

                    if ((isProjectMatched && isProjectCollectionMatched && isServerMatched) || (documentModel.MappingShowName != null && documentModel.MappingShowName.Equals(template.ShowName)))
                    {
                        var rddi = Factory.CreateRibbonDropDownItem();
                        rddi.Label = template.ShowName;
                        rddi.SuperTip = string.Format(CultureInfo.CurrentCulture, Resources.TM_ToolTip, template.ShowName, Path.GetFileName(template.TemplateFile));
                        if (File.Exists(template.TemplateFavicon))
                        {
                            // TempFolder creates only one copy of this file regardles how many time is this method called.
                            var copyFile = TempFolder.CreateTemporaryFile(template.TemplateFavicon);
                            if (!string.IsNullOrEmpty(copyFile) && File.Exists(copyFile))
                                rddi.Image = Image.FromFile(copyFile);
                        }

                        dropDownSelectTemplate.Items.Add(rddi);
                    }
                }
            }

            dropDownSelectTemplate.Enabled = dropDownSelectTemplate.Items.Count > 0;
            menuNewEmptyWorkItem.Enabled = dropDownSelectTemplate.Items.Count > 0;
            menuNewWorkItem.Enabled = dropDownSelectTemplate.Items.Count > 0;
        }

        /// <summary>
        /// Method inserts the requirement table on the actual position in actual document.
        /// </summary>
        /// <returns>Returns the inserted table. Null if no table inserted.</returns>
        private static Table InsertReqTable(ISyncServiceDocumentModel documentModel, IConfigurationItem item)
        {
            var document = documentModel.WordDocument as Document;
            if (document == null)
            {
                return null;
            }

            var xmlFile = item.RelatedTemplateFile;
            SyncServiceTrace.I(Resources.LogService_InsertTableFromXmlFile, xmlFile);

            if (!File.Exists(xmlFile))
            {
                SyncServiceTrace.W(Resources.LogService_InsertTableFromXmlFile_FileDoesNotExist, xmlFile);
                return null;
            }

            var selection = document.ActiveWindow.Selection;


            if (WordSyncHelper.IsCursorBehindTable(selection))
            {
                selection.TypeParagraph();
            }

            if (WordSyncHelper.IsCursorInTable(selection))
            {
                SyncServiceTrace.W(Resources.WordToTFS_Error_TableInsertInTable);
                MessageBox.Show(Resources.WordToTFS_Error_TableInsertInTable,
                                Resources.WordToTFS_Error_Title,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                return null;
            }

            selection.Range.InsertFile(FileName: xmlFile, ConfirmConversions: false);

            var tables = selection.Tables;
            if (tables.Count != 1)
            {
                SyncServiceTrace.W(Resources.LogService_InsertTableFromXmlFile_Failed, xmlFile);
                return null;
            }

            return tables.Cast<Table>().FirstOrDefault();
        }

        /// <summary>
        /// Method attaches the model to the ribbon.
        /// </summary>
        /// <param name="model"><see cref="ISyncServiceModel">Model</see> to attach.</param>
        public void AttachModel(ISyncServiceModel model)
        {
            if (WordApplicationModel != null)
            {
                WordApplicationModel.ActiveDocumentChanged -= HandleModelActiveDocumentChanged;
                WordApplicationModel.DocumentOpen -= HandleModelDocumentOpen;
            }
            WordApplicationModel = model;
            SyncServiceFactory.RegisterService(WordApplicationModel);
            if (WordApplicationModel != null)
            {
                WordApplicationModel.ActiveDocumentChanged += HandleModelActiveDocumentChanged;
                WordApplicationModel.DocumentOpen += HandleModelDocumentOpen;
            }
        }

        /// <summary>
        /// Adds one <see cref="RibbonButton"/> for each work item in the template files to the <see cref="RibbonSplitButton"/>.
        /// </summary>
        /// <param name="documentModel">Model of associated word document.</param>
        /// <param name="menu"><see cref="RibbonMenu"/> to be filled.</param>
        /// <param name="labelPrefix">A string that is added to the work item label.</param>
        /// <param name="listButtonClickEvent">The <see cref="EventHandler{TEventArgs}"/> the <see cref="RibbonButton"/>s shall react to.</param>
        private void LoadWorkItemDropdownMenu(ISyncServiceDocumentModel documentModel, RibbonMenu menu, string labelPrefix, RibbonControlEventHandler listButtonClickEvent)
        {
            if (dropDownSelectTemplate.SelectedItem == null)
            {
                return;
            }
            menu.Items.Clear();

            foreach (var configurationItem in documentModel.Configuration.GetConfigurationItems())
            {
                // Check if the element should be shown in word.
                if (!configurationItem.HideElementInWord)
                {
                    var ribbonButton = Factory.CreateRibbonButton();

                    ribbonButton.Tag = configurationItem;
                    ribbonButton.Label = labelPrefix + configurationItem.WorkItemTypeMapping;
                    ribbonButton.Image = configurationItem.ImageFile;
                    ribbonButton.Click += listButtonClickEvent;

                    menu.Items.Add(ribbonButton);
                }
            }
        }

        /// <summary>
        /// Called if the visible property in pane is changed.
        /// </summary>
        /// <param name="control">Control changed visible property.</param>
        /// <param name="show">How the property is changed.</param>
        internal void VisibleChanged(UserControl control, bool show)
        {
            if (!show)
            {
                UncheckButton(control);
            }

            // This looks fugly, but I also didn't want to use 'dynamic'
            IHostPaneControl hostedControl = null;
            if (control is HostPane<TestSpecificationReport>)
            {
                hostedControl = (control as HostPane<TestSpecificationReport>).HostedControl;
            }
            else if (control is HostPane<TestSpecificationReportByQuery>)
            {
                hostedControl = (control as HostPane<TestSpecificationReportByQuery>).HostedControl;
            }
            else if (control is HostPane<TestResultReport>)
            {
                hostedControl = (control as HostPane<TestResultReport>).HostedControl;
            }
            else if (control is HostPane<SynchronizationStatePanelView>)
            {
                hostedControl = (control as HostPane<SynchronizationStatePanelView>).HostedControl;
            }
            else if (control is HostPane<GetWorkItemsPanelView>)
            {
                hostedControl = (control as HostPane<GetWorkItemsPanelView>).HostedControl;
            }

            // hosted control may be null if panel was hidden by code
            if (hostedControl != null)
            {
                hostedControl.VisibilityChanged(show);
            }
        }

        /// <summary>
        /// The methods sets the check box off.
        /// </summary>
        /// <param name="control">Control determines which check box should be switched off.</param>
        internal void UncheckButton(UserControl control)
        {
            if (control is HostPane<TemplateManagerControl>)
            {
                buttonTemplateManager.Checked = false;
            }
            else if (control is HostPane<DefaultValuesControl>)
            {
                buttonEditDefVals.Checked = false;
            }
            else if (control is HostPane<GetWorkItemsPanelView>)
            {
                buttonGetWorkItems.Checked = false;
            }
            else if (control is HostPane<SynchronizationStatePanelView>)
            {
                buttonSynchronizationState.Checked = false;
            }
            else if (control is HostPane<AreaIterationPathView>)
            {
                buttonAreaIterationPath.Checked = false;
            }
            else if (control is HostPane<TestSpecificationReport>)
            {
                splitButtonTestSpecificationReport.Checked = false;
            }
            else if (control is HostPane<TestSpecificationReportByQuery>)
            {
                buttonTestSpecificationReportByQuery.Checked = false;
            }
            else if (control is HostPane<TestResultReport>)
            {
                buttonTestResultReport.Checked = false;
            }
        }

        /// <summary>
        /// Show/Hide ribbon controls based on config settings.
        /// </summary>
        private void EnableRibbonControls(ISyncServiceDocumentModel documentModel)
        {
            // Set default values
            var publishEnabled = true;
            var refreshEnabled = false;
            var getWorkItemsEnabled = true;
            var synchronizationStateEnabled = true;
            var ignoreFormattingEnabled = true;
            var ignoreConflictsEnabled = true;
            var emptyEnabled = true;
            var newEnabled = true;
            var showHeader = false;
            var editDefaultValuesEnabled = true;
            var deleteIdsEnabled = true;
            var areaIterationPathEnabled = true;
            var testSpecificationAvailable = false;
            var testResultAvailable = false;
            var templateManagerEnabled = true;
            var templateSelectionEnabled = true;

            if (documentModel.Configuration != null)
            {
                // Read the configuration from w2t file.
                publishEnabled = documentModel.Configuration.EnablePublish;
                refreshEnabled = documentModel.Configuration.EnableRefresh;
                getWorkItemsEnabled = documentModel.Configuration.EnableGetWorkItems;
                synchronizationStateEnabled = documentModel.Configuration.EnableOverview;
                ignoreFormattingEnabled = documentModel.Configuration.EnableSwitchFormatting;
                ignoreConflictsEnabled = documentModel.Configuration.EnableSwitchConflictOverwrite;
                emptyEnabled = documentModel.Configuration.EnableEmpty;
                newEnabled = documentModel.Configuration.EnableNew;
                showHeader = documentModel.Configuration.Headers.Count > 0;
                editDefaultValuesEnabled = documentModel.Configuration.EnableEditDefaultValues;
                deleteIdsEnabled = documentModel.Configuration.EnableDeleteIds;
                areaIterationPathEnabled = documentModel.Configuration.EnableAreaIterationPath;
                testSpecificationAvailable = documentModel.Configuration.ConfigurationTest.ConfigurationTestSpecification.Available;
                testResultAvailable = documentModel.Configuration.ConfigurationTest.ConfigurationTestResult.Available;
                templateManagerEnabled = documentModel.Configuration.EnableTemplateManager;
                templateSelectionEnabled = documentModel.Configuration.EnableTemplateSelection;

                var buttons = documentModel.Configuration.Buttons;
                if (buttons != null && buttons.Any())
                {
                    foreach (var button in buttons)
                    {
                        switch (button.Name)
                        {
                            case "Publish":
                                splitButtonPublish.Label = button.Text;
                                break;
                            case "Refresh":
                                splitButtonRefresh.Label = button.Text;
                                break;
                            case "Get Work Items":
                                buttonGetWorkItems.Label = button.Text;
                                break;
                            case "Compare":
                                buttonSynchronizationState.Label = button.Text;
                                break;
                            case "Empty Work Item":
                                menuNewEmptyWorkItem.Label = button.Text;
                                break;
                            case "New Work Item":
                                menuNewWorkItem.Label = button.Text;
                                break;
                            case "Insert Header":
                                menuNewHeader.Label = button.Text;
                                break;
                            case "Edit Default Values":
                                buttonEditDefVals.Label = button.Text;
                                break;
                            case "Delete Ids":
                                buttonDeleteIds.Label = button.Text;
                                break;
                            case "Area und Iteration Path":
                                buttonAreaIterationPath.Label = button.Text;
                                break;
                            case "Test Spec Report":
                                splitButtonTestSpecificationReport.Label = button.Text;
                                break;
                            case "Test Spec Report By Query":
                                buttonTestSpecificationReportByQuery.Label = button.Text;
                                break;
                            case "Test Result Report":
                                buttonTestResultReport.Label = button.Text;
                                break;
                            default:
                                System.Windows.MessageBox.Show("The button with the name \"" + button.Name + "\" does not exist. Please check your template.", "Warning!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                break;
                        }
                    }
                }
                else
                {
                    splitButtonPublish.Label = Resources.ButtonPublish;
                    splitButtonRefresh.Label = Resources.ButtonRefresh;
                    buttonGetWorkItems.Label = Resources.ButtonGetWorkItems;
                    buttonSynchronizationState.Label = Resources.ButtonCompare;
                    menuNewEmptyWorkItem.Label = Resources.ButtonEmpyWorkItem;
                    menuNewWorkItem.Label = Resources.ButtonNewWorkItem;
                    menuNewHeader.Label = Resources.ButtonHeader;
                    buttonEditDefVals.Label = Resources.ButtonDefaultValues;
                    buttonDeleteIds.Label = Resources.ButtonDeleteIds;
                    buttonAreaIterationPath.Label = Resources.ButtonAreaAndIterationPath;
                    splitButtonTestSpecificationReport.Label = Resources.ButtonTestSpecReport;
                    buttonTestSpecificationReportByQuery.Label = Resources.ButtonTestSpecReportByQuery;
                    buttonTestResultReport.Label = Resources.ButtonTestResultReport;
                    buttonSettings.Label = Resources.ButtonSettings;
                }
            }

            // Publish group
            splitButtonPublish.Visible = publishEnabled;
            splitButtonRefresh.Visible = refreshEnabled;
            buttonGetWorkItems.Visible = getWorkItemsEnabled;
            buttonSynchronizationState.Visible = synchronizationStateEnabled;

            // Template group
            chkIgnoreFormatting.Visible = ignoreFormattingEnabled;
            chkConflictOverwrite.Visible = ignoreConflictsEnabled;
            buttonTemplateManager.Visible = templateManagerEnabled;
            dropDownSelectTemplate.Enabled = templateSelectionEnabled;

            // Work items group
            menuNewEmptyWorkItem.Visible = emptyEnabled;
            menuNewWorkItem.Visible = newEnabled;
            menuNewHeader.Visible = showHeader;
            buttonEditDefVals.Visible = editDefaultValuesEnabled;
            buttonDeleteIds.Visible = deleteIdsEnabled;
            buttonAreaIterationPath.Visible = areaIterationPathEnabled;
            groupRequirements.Visible = emptyEnabled | newEnabled | editDefaultValuesEnabled | deleteIdsEnabled | areaIterationPathEnabled;

            // Test management group
            splitButtonTestSpecificationReport.Visible = testSpecificationAvailable;
            splitButtonTestSpecificationReport.Enabled = documentModel.TfsDocumentBound;
            buttonTestResultReport.Visible = testResultAvailable;
            buttonTestResultReport.Enabled = documentModel.TfsDocumentBound;
            groupTestManagement.Visible = splitButtonTestSpecificationReport.Visible || buttonTestResultReport.Visible;
        }

        /// <summary>
        /// Initialize and validate Refresh button process.
        /// </summary>
        /// <param name="destinationIds">Ids of the work items to refresh</param>
        /// <param name="source">TFS adapter.</param>
        /// <param name="sourceWorkItems">TFS work items to import.</param>
        /// <returns>If true then validation succeeded, false validation failed.</returns>
        private bool ValidateRefresh(int[] destinationIds, IWorkItemSyncAdapter source, out IList<IWorkItem> sourceWorkItems)
        {
            sourceWorkItems = null;

            if (source == null || destinationIds == null || destinationIds.Length == 0)
            {
                return false;
            }

            // get only unique ids
            destinationIds = destinationIds.Distinct().ToArray();

            // open TFS adapter with Word document work item ids
            if (!source.Open(destinationIds))
            {
                return false;
            }

            // for sure filter out TFS WorkItems contained also in Word document
            sourceWorkItems = source.WorkItems.Where(wi => destinationIds.Contains(wi.Id)).ToList();
            // if no any then return
            if (sourceWorkItems.Count == 0)
            {
                MessageBox.Show(this, Resources.Refresh_NoIdentifiedWorkItems, Resources.RefreshWorkItemsCaption, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
                return false;
            }

            // if the count of word document work items is not equal to count of TFS work items
            if (destinationIds.Length != sourceWorkItems.Count)
            {
                MessageBox.Show(this, Resources.Refresh_SomeNoIdentifiedWorkItems, Resources.RefreshWorkItemsCaption, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
                return false;
            }
            return true;
        }

        #endregion Private methods

        #region Private event handler methods

        /// <summary>
        /// Called when a document is opened.
        /// </summary>
        private void HandleModelDocumentOpen(object sender, DocumentOpenEventArgs e)
        {
            if (Globals.ThisAddIn.Application.Documents.Count == 0)
            {
                return;
            }

            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            documentModel.Configuration.RefreshMappings();
            documentModel.Configuration.ActivateAllMappings();

            // In case document was saved with varible set to true we have to reset it.
            documentModel.TestReportRunning = false;

            // Check if the required template (mapping) exists.
            // Show message if:
            //   - Template (mapping) defined
            //   - Template (mapping) not exists in templates (mappings)
            //   - Document is bound
            if (documentModel.TfsDocumentBound && !string.IsNullOrEmpty(documentModel.MappingShowName) && !documentModel.Configuration.MappingExists(documentModel.MappingShowName))
            {
                var text = string.Format(CultureInfo.CurrentCulture, Resources.Error_DocumentTemplateNotExists, documentModel.MappingShowName);
                MessageBox.Show(this, text, Resources.Error_MessageBoxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, 0);
                documentModel.MappingShowName = null;
                documentModel.Save();
            }

            UpdateTemplateSelection(documentModel);
            EnableAllControls(documentModel);

            //Automatic Binding
            TryBindProject();

            //Automatic Refresh
            TryAutomaticRefresh();

        }

        /// <summary>
        /// Bind the project to the tfs if the autoconnect is specified in the template
        /// </summary>
        private void TryBindProject()
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);

            if (documentModel.Configuration.AutoConnect)
            {

                if (documentModel.TfsDocumentBound)
                {
                    documentModel.UnbindProject();
                    SetHostPaneVisibility<SynchronizationStatePanelView>(documentModel.WordDocument as Document, false);
                    SetHostPaneVisibility<GetWorkItemsPanelView>(documentModel.WordDocument as Document, false);
                    SetHostPaneVisibility<AreaIterationPathView>(documentModel.WordDocument as Document, false);
                    SetHostPaneVisibility<TestSpecificationReport>(documentModel.WordDocument as Document, false);
                    SetHostPaneVisibility<TestSpecificationReportByQuery>(documentModel.WordDocument as Document, false);
                    SetHostPaneVisibility<TestResultReport>(documentModel.WordDocument as Document, false);
                }

                    documentModel.BindProject();
            }

            EnableAllControls(documentModel);
        }

        /// <summary>
        /// Try an automatic refresh by a query
        /// </summary>
        private void TryAutomaticRefresh()
        {

            //Prepare the document Model
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);

            if (string.IsNullOrEmpty(documentModel.Configuration.AutoRefreshQuery))
            {
                return;
            }

            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.NewProgress(Resources.RefreshWorkItemsCaptionAutomatic);
            progressService.EnterProgressGroup(2, Resources.RefreshWorkItemsAutomaticPreparationText);
            progressService.ShowProgress();


            //Prepare the Query
            var ids = GetWorkItemsIdsByQueryName(documentModel.Configuration.AutoRefreshQuery, documentModel);
            var qc = new QueryConfiguration();

            foreach (var workItem in ids)
            {
                qc.ByIDs.Add(workItem);
            }
            qc.ImportOption = QueryImportOption.IDs;

            var source = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(documentModel.TfsServer, documentModel.TfsProject, qc, documentModel.Configuration);

            // start background thread to execute the import
            var backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += DoPreparationForAutoRefresh;
            backgroundWorker.RunWorkerCompleted += (s, e) =>
                                                       {
                                                           backgroundWorker.Dispose();
                                                           PreparationFinished(documentModel, source, e);
                                                       };
            backgroundWorker.RunWorkerAsync(source);

        }

        /// <summary>
        /// Occurs when the background operation has completed, has been canceled, or has raised an exception.
        /// </summary>
        private void PreparationFinished(ISyncServiceDocumentModel documentModel, IWorkItemSyncAdapter source, RunWorkerCompletedEventArgs e)
        {

            e.Error.NotifyIfException("Failed to refresh work items.");

            var progressService = SyncServiceFactory.GetService<IProgressService>();

            progressService.EnterProgressGroup(2, Resources.RefreshWorkItemsAutomaticPreparationText);
            if (progressService.ProgressCanceled)
            {
                progressService.HideProgress();
            }
            else
            {
                InitiateRefresh(documentModel, source.WorkItems);
            }
        }

        /// <summary>
        /// Background method to prepare the adapter for the import. The open method takes some time, it is therefore implemented as a Background worker
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e"></param>
        private void DoPreparationForAutoRefresh(object sender, DoWorkEventArgs e)
        {
            var source = (IWorkItemSyncAdapter)e.Argument;
            source.Open(null);

        }

        /// <summary>
        /// Get all work item ids for a given query name
        /// </summary>
        /// <param name="query">The name of the query</param>
        /// <param name="documentModel">The current document model</param>
        /// <returns></returns>
        private static IEnumerable<int> GetWorkItemsIdsByQueryName(string query, ISyncServiceDocumentModel documentModel)
        {
            var queryconfig = new QueryConfiguration
            {
                QueryPath = query,
                ImportOption = QueryImportOption.SavedQuery
            };
            var adapter = new Tfs2012SyncAdapter(documentModel.TfsServer, documentModel.TfsProject, queryconfig, documentModel.Configuration);
            var intIDs = new List<int>();

            IList<WorkItem> queryWorkItems;
            //try
            //{
            queryWorkItems = adapter.LoadWorkItemsFromSavedQuery(queryconfig);

            //}
            //catch (Exception)
            //{
            //    //PublishUserInformation("Query does not exists", "Query does not exsist");
            //    return intIDs;
            //}

            foreach (var queryWorkItem in queryWorkItems)
            {
                intIDs.Add(queryWorkItem.Id);
            }

            return intIDs;
        }

        /// <summary>
        /// Called when the active document changed. No document may be active if no document opened.
        /// </summary>
        private void HandleModelActiveDocumentChanged(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Documents.Count == 0)
            {
                return;
            }

            if (Globals.ThisAddIn.Application.Documents.Count == 1)
            {
                // When opening a document using doublclick in explorer, no DocumentOpen event is raised...
                HandleModelDocumentOpen(sender, null);
            }

            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            // refresh templates, but make sure the ignore formatting and overwrite settings are not reset
            var ignoreFormatting = documentModel.Configuration.IgnoreFormatting;
            var conflictOverwrite = documentModel.Configuration.ConflictOverwrite;
            documentModel.Configuration.RefreshMappings();
            documentModel.Configuration.ActivateAllMappings();

            UpdateTemplateSelection(documentModel);

            documentModel.Configuration.IgnoreFormatting = ignoreFormatting;
            documentModel.Configuration.ConflictOverwrite = conflictOverwrite;
            chkIgnoreFormatting.Checked = ignoreFormatting;
            chkConflictOverwrite.Checked = conflictOverwrite;

            EnableAllControls(documentModel);
        }

        /// <summary>
        /// Called when the user clicks the button 'buttonBindUnbind'.
        /// </summary>
        private void HandleButtonBindUnbindClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            if (documentModel.TfsDocumentBound)
            {
                documentModel.UnbindProject();
                SetHostPaneVisibility<SynchronizationStatePanelView>(documentModel.WordDocument as Document, false);
                SetHostPaneVisibility<GetWorkItemsPanelView>(documentModel.WordDocument as Document, false);
                SetHostPaneVisibility<AreaIterationPathView>(documentModel.WordDocument as Document, false);
                SetHostPaneVisibility<TestSpecificationReport>(documentModel.WordDocument as Document, false);
                SetHostPaneVisibility<TestSpecificationReportByQuery>(documentModel.WordDocument as Document, false);
                SetHostPaneVisibility<TestResultReport>(documentModel.WordDocument as Document, false);
            }
            else
            {
                documentModel.BindProject();
                Cursor.Current = Cursors.WaitCursor;
                var missingFields = documentModel.GetMissingFields();
                if (missingFields.Count > 0)
                {
                    var informationWindow = new MissingFieldsInformationWindow
                                                {
                                                    DataContext = missingFields
                                                };
                    informationWindow.ShowDialog();
                }

                var wronglyMappedField = documentModel.GetWronglyMappedFields();
                if (wronglyMappedField.Count > 0)
                {
                    var wronglyMappedFieldsinformationWindow = new WronglyMappedFieldsInformationWindow()
                    {
                        DataContext = wronglyMappedField
                    };
                    wronglyMappedFieldsinformationWindow.ShowDialog();
                }
                Cursor.Current = Cursors.Default;
                TryAutomaticRefresh();
            }
            UpdateTemplateSelection(documentModel);
            EnableAllControls(documentModel);
        }

        /// <summary>
        /// Called when the user clicks the button 'buttonPublish' or 'buttonForcePublish'.
        /// </summary>
        private void HandleButtonPublishClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                Debug.Assert(documentModel == null, "Document model not found!");
                return;
            }

            // Reset old showed messages.
            ResetBeforeOperation(documentModel);

            var sourceAdapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(documentModel.WordDocument, documentModel.Configuration);

            if (!sourceAdapter.Open(null))
            {
                return;
            }

            if (sourceAdapter.WorkItems == null || sourceAdapter.WorkItems.Count == 0)
            {
                MessageBox.Show(this, Resources.Error_NoWorkItemsToPublish, Resources.PublishTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
            }
            else
            {
                InitiatePublish(documentModel, sourceAdapter.WorkItems);
            }
        }

        /// <summary>
        /// Called when the user clicks the button 'buttonPublishSelected'.
        /// </summary>
        private void HandleButtonPublishSelectedClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                Debug.Assert(documentModel == null, "Document model not found!");
                return;
            }

            // Reset old showed messages.
            ResetBeforeOperation(documentModel);

            var sourceAdapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(documentModel.WordDocument, documentModel.Configuration);

            if (!sourceAdapter.Open(null))
            {
                return;
            }

            if (!sourceAdapter.GetSelectedWorkItems().Any())
            {
                MessageBox.Show(this, Resources.Error_NoWorkItemSelected, Resources.PublishTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
            }
            else
            {
                InitiatePublish(documentModel, sourceAdapter.GetSelectedWorkItems());
            }
        }

        /// <summary>
        /// Called when the user clicks the button 'buttonDeleteIds'.
        /// </summary>
        private void HandleButtonDeleteIdsClick(object sender, RibbonControlEventArgs eventArgs)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            if (DialogResult.Yes ==
                MessageBox.Show(Resources.WordToTFS_IdDeletionDialog, Resources.WordToTFS_IdDeletionCation, MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, 0))
            {
                var wordAdapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(documentModel.WordDocument, documentModel.Configuration);

                wordAdapter.Open(null);

                foreach (var workitem in wordAdapter.WorkItems)
                {
                    workitem.DeleteIdAndRevision();
                }
            }
        }

        /// <summary>
        /// Open help document.
        /// </summary>
        private void HandleButtonHelpClick(object sender, RibbonControlEventArgs e)
        {
            var process = Process.Start(Resources.WordToTFS_Help);
            if (process != null)
            {
                process.Dispose();
            }
        }

        /// <summary>
        /// Opens the about dialog if the user clicks the button. Saves (debug) settings afterwards.
        /// </summary>
        private void HandleButtonSettingsClick(object sender, RibbonControlEventArgs e)
        {
            var wnd = new SettingsWindow();
            (new WindowInteropHelper(wnd)).Owner = Handle;
            wnd.ShowDialog();
            var level = ((SettingsModel)(wnd.DataContext)).SelectedLevel;
            Settings.Default.EnableDebugging = (int)level;
            Settings.Default.Save();
        }

        /// <summary>
        /// Opens the about settings if the user clicks the button. Saves settings afterwards.
        /// </summary>
        private void HandleButtonAboutClick(object sender, RibbonControlEventArgs e)
        {
            var wnd = new AboutWindow();
            (new WindowInteropHelper(wnd)).Owner = Handle;
            wnd.ShowDialog();
            Settings.Default.Save();
        }

        /// <summary>
        /// Called when the user clicks the ignore formatting checkbox
        /// </summary>
        private void HandleCheckboxIgnoreFormattingClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null || sender is RibbonCheckBox == false)
            {
                return;
            }

            documentModel.Configuration.IgnoreFormatting = ((RibbonCheckBox)sender).Checked;
        }

        /// <summary>
        /// Called when the user clicks a button in the list of <see cref="menuNewEmptyWorkItem"/>.
        /// </summary>
        private void HandleMenuNewEmptyWorkItemClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null || sender is RibbonButton == false)
            {
                return;
            }

            var item = ((RibbonButton)sender).Tag as IConfigurationItem;
            if (item == null)
            {
                return;
            }
            documentModel.Save();
            InsertReqTable(documentModel, item);
        }

        /// <summary>
        /// Called when the user clicks a button in the list of <see cref="menuNewWorkItem"/>.
        /// </summary>
        private void HandleMenuNewWorkItemClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null || sender is RibbonButton == false)
            {
                return;
            }

            var item = ((RibbonButton)sender).Tag as IConfigurationItem;
            if (item == null)
            {
                return;
            }

            var table = InsertReqTable(documentModel, item);
            if (table != null)
            {
                documentModel.Save();
                var adapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(documentModel.WordDocument, documentModel.Configuration);
                var newItem = adapter.GetSelectedWorkItems().First();

                System.Threading.Tasks.Task.Factory.StartNew(() =>
                                                                 {
                                                                     LoadAllowedValues(documentModel, newItem, false);
                                                                     SetDefaultValues(item, documentModel, newItem);
                                                                 });
            }
        }

        /// <summary>
        /// Called when the user clicks a button in the list of <see cref="menuNewHeader"/>.
        /// </summary>
        private void HandleMenuNewHeaderClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null || sender is RibbonButton == false)
            {
                return;
            }

            var item = ((RibbonButton)sender).Tag as IConfigurationItem;
            if (item == null)
            {
                return;
            }

            var table = InsertReqTable(documentModel, item);
            if (table != null)
            {
                documentModel.Save();
                var adapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(documentModel.WordDocument, documentModel.Configuration);
                var newItem = adapter.GetSelectedHeader().First();
                System.Threading.Tasks.Task.Factory.StartNew(
                    () =>
                    {
                        LoadAllowedValues(documentModel, newItem, true);
                        SetDefaultValues(item, documentModel, newItem);
                    });
            }
        }

        /// <summary>
        /// Occurs when a user selects a new item on a Ribbon drop-down control.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event data.</param>
        private void HandleDropDownSelectTemplateSelectionChanged(object sender, RibbonControlEventArgs e)
        {
            SyncServiceTrace.D("Selecting new template {0}", dropDownSelectTemplate.SelectedItem.Label);
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            SetHostPaneVisibility<TestSpecificationReport>(documentModel.WordDocument as Document, false);
            SetHostPaneVisibility<TestSpecificationReportByQuery>(documentModel.WordDocument as Document, false);
            SetHostPaneVisibility<TestResultReport>(documentModel.WordDocument as Document, false);

            Cursor.Current = Cursors.WaitCursor;
            documentModel.Configuration.ActivateMapping(dropDownSelectTemplate.SelectedItem.Label);


            ActivateTemplate(documentModel, dropDownSelectTemplate.SelectedItem.Label);


            //Try the automatic bind
            TryBindProject();

            var missingFields = documentModel.GetMissingFields();
            if (missingFields.Count > 0)
            {
                var informationWindow = new MissingFieldsInformationWindow
                {
                    DataContext = missingFields
                };
                informationWindow.ShowDialog();
            }

            var wronglyMappedField = documentModel.GetWronglyMappedFields();
            if (wronglyMappedField.Count > 0)
            {
                var wronglyMappedFieldsinformationWindow = new WronglyMappedFieldsInformationWindow()
                {
                    DataContext = wronglyMappedField
                };
                wronglyMappedFieldsinformationWindow.ShowDialog();
            }

            Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// Handles the click event of the feedback button.
        /// </summary>
        private void HandleFeedbackButtonClick(object sender, RibbonControlEventArgs e)
        {
            var process = Process.Start(Resources.WordToTFS_FeedbackWebPage);
            if (process != null)
            {
                process.Dispose();
            }
        }

        /// <summary>
        /// Handles the click event of the Update button
        /// The method will determine the url from the registry and will try to update it from this path
        /// If WordToTFS was installed from a network drive or similar --> The version will be updated from this path.
        /// If the method fails, the standard path from the ressources will be used
        /// </summary>
        private void HandleButtonUpdateClick(object sender, RibbonControlEventArgs e)
        {
            string updatePathFromRegistry = null;
            if (ApplicationDeployment.CurrentDeployment != null)
            {
                updatePathFromRegistry = ApplicationDeployment.CurrentDeployment.UpdateLocation.ToString();
            }

            if (string.IsNullOrWhiteSpace(updatePathFromRegistry))
            {
                var process = Process.Start(Resources.WordToTFS_Update);
                if (process != null)
                {
                    process.Dispose();
                }
            }
            else
            {
                var process = Process.Start(updatePathFromRegistry);
                if (process != null)
                {
                    process.Dispose();
                }
            }
        }

        /// <summary>
        /// Called when the user clicks the button 'Refresh selected work item'.
        /// </summary>
        private void ButtonRefreshSelectedClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            // Reset old showed messages.
            ResetBeforeOperation(documentModel);

            var destination = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(documentModel.WordDocument, documentModel.Configuration);

            if (!destination.Open(null))
            {
                return;
            }

            if (!destination.GetSelectedWorkItems().Any())
            {
                MessageBox.Show(this, Resources.Error_NoWorkItemSelected, Resources.RefreshWorkItemsCaption, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
            }
            else
            {
                // ask whether to overide local changes in WorkItems in Word
                var message = string.Format(CultureInfo.CurrentCulture,
                                            Resources.TFSRefreshSelected_LocalChanges,
                                            string.Join(", ", destination.GetSelectedWorkItems().Select(x => x.Id.ToString(CultureInfo.InvariantCulture)).ToArray()));

                SyncServiceTrace.W(Resources.TFSRefresh_LocalChanges);
                if (MessageBox.Show(this, message, Resources.WordToTFS_Error_Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, 0) == DialogResult.No)
                {
                    return;
                }

                InitiateRefresh(documentModel, destination.GetSelectedWorkItems());
            }
        }

        /// <summary>
        /// Called when the user clicks the button 'Refresh'.
        /// </summary>
        private void ButtonRefreshClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            // Reset old showed messages.
            ResetBeforeOperation(documentModel);

            var destination = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(documentModel.WordDocument, documentModel.Configuration);

            if (!destination.Open(null))
            {
                return;
            }

            if (destination.WorkItems == null || !destination.WorkItems.Any())
            {
                MessageBox.Show(this, Resources.Error_NoWorkItemsToRefresh, Resources.RefreshWorkItemsCaption, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
            }
            else
            {
                // ask whether to overide local changes in WorkItems in Word
                SyncServiceTrace.W(Resources.TFSRefresh_LocalChanges);
                if (MessageBox.Show(this, Resources.TFSRefresh_LocalChanges, Resources.WordToTFS_Error_Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, 0) ==
                    DialogResult.No)
                {
                    return;
                }

                InitiateRefresh(documentModel, destination.WorkItems);
            }
        }

        #endregion Private event handler methods

        #region Show and hide task panes

        /// <summary>
        /// Handles the button template manager click.
        /// </summary>
        private void HandleButtonTemplateManagerClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            // Show or hide pane depending on button checked state
            TemplateManagerControl taskPane = null;
            if (buttonTemplateManager.Checked)
            {
                taskPane = new TemplateManagerControl(documentModel, _templateManager);
            }

            SetHostPaneVisibility(documentModel.WordDocument as Document, buttonTemplateManager.Checked, taskPane, Resources.TemplateManagerPanelCaption);
        }

        /// <summary>
        /// Shows the pane to edit default values when the user clicks the ribbon button.
        /// </summary>
        private void HandleButtonEditDefValsClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            SetHostPaneVisibility(documentModel.WordDocument as Document, buttonEditDefVals.Checked, new DefaultValuesControl(documentModel), Resources.DefaultValuesPanelCaption);
        }

        /// <summary>
        /// Called when the user clicks the button 'buttonGetWorkItems'.
        /// </summary>
        private void HandleButtonGetWorkItemsClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            GetWorkItemsPanelView paneContent = null;
            if (buttonGetWorkItems.Checked)
            {
                var tfsService = SyncServiceFactory.CreateTfsService(documentModel.TfsServer, documentModel.TfsProject, documentModel.Configuration);
                paneContent = new GetWorkItemsPanelView
                                  {
                                      Model = new GetWorkItemsPanelViewModel(tfsService, documentModel, this)
                                  };
            }

            SetHostPaneVisibility(documentModel.WordDocument as Document, buttonGetWorkItems.Checked, paneContent, Resources.GetWorkItemsPanelCaption);
        }

        /// <summary>
        /// Called when the user clicks the button 'buttonSynchronizationState'.
        /// </summary>
        private void HandleButtonSynchronizationStateClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            SynchronizationStatePanelView paneContent = null;
            if (buttonSynchronizationState.Checked)
            {
                var tfsService = SyncServiceFactory.CreateTfsService(documentModel.TfsServer, documentModel.TfsProject, documentModel.Configuration);
                paneContent = new SynchronizationStatePanelView
                {
                    Model = new WorkItemOverviewPanelViewModel(tfsService, documentModel, this)
                };
            }

            SetHostPaneVisibility(documentModel.WordDocument as Document, buttonSynchronizationState.Checked, paneContent, Resources.SynchronizationStatePanelCaption);
        }

        /// <summary>
        /// Shows or hides the test specification report task pane
        /// </summary>
        private void HandleButtonTestSpecificationReportClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }
            documentModel.Save();

            // If the button is disabled in the configuration, dont show pane and disable button
            if (splitButtonTestSpecificationReport.Checked && !documentModel.Configuration.ConfigurationTest.ConfigurationTestSpecification.Available)
            {
                splitButtonTestSpecificationReport.Enabled = false;
                splitButtonTestSpecificationReport.Checked = false;
            }

            // Set visibility of pane to checked state of button
            TestSpecificationReport paneContent = null;
            if (splitButtonTestSpecificationReport.Checked)
            {
                var testAdapter = SyncServiceFactory.CreateTfsTestAdapter(documentModel.TfsServer, documentModel.TfsProject, documentModel.Configuration);
                paneContent = new TestSpecificationReport(documentModel.WordDocument as Document);
                var wordToTfsViewDispatcher = new ViewDispatcher(paneContent.Dispatcher);
                paneContent.Model = new TestSpecificationReportModel(documentModel, wordToTfsViewDispatcher, testAdapter, this, new TestReportingProgressCancellationService(true));
            }
            SetHostPaneVisibility(documentModel.WordDocument as Document, splitButtonTestSpecificationReport.Checked, paneContent, Resources.TestSpecificationReportPanelCaption);
        }

        /// <summary>
        /// Handles the click event of the test specification report by query button.
        /// </summary>
        /// <param name="sender">Sender of event.</param>
        /// <param name="e">Event data.</param>
        private void HandleButtonTestSpecificationReportByQueryClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }
            documentModel.Save();

            // If the button is disabled in the configuration, dont show pane and disable button
            if (buttonTestSpecificationReportByQuery.Checked && !documentModel.Configuration.ConfigurationTest.ConfigurationTestSpecification.Available)
            {
                buttonTestSpecificationReportByQuery.Enabled = false;
                buttonTestSpecificationReportByQuery.Checked = false;
            }

            // Set visibility of pane to checked state of button
            TestSpecificationReportByQuery paneContent = null;
            if (buttonTestSpecificationReportByQuery.Checked)
            {
                var testAdapter = SyncServiceFactory.CreateTfsTestAdapter(documentModel.TfsServer, documentModel.TfsProject, documentModel.Configuration);
                var tfsService = SyncServiceFactory.CreateTfsService(documentModel.TfsServer, documentModel.TfsProject, documentModel.Configuration);

                paneContent = new TestSpecificationReportByQuery
                {
                    Model = new TestSpecificationReportByQueryModel(tfsService, documentModel, this, testAdapter)
                };
            }
            SetHostPaneVisibility(documentModel.WordDocument as Document, buttonTestSpecificationReportByQuery.Checked, paneContent, Resources.TestSpecificationReportByQueryPanelCaption);
        }

        /// <summary>
        /// Handles the click event of the test result report button.
        /// </summary>
        /// <param name="sender">Sender of event.</param>
        /// <param name="e">Event data.</param>
        private void HandleButtonTestResultReportClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }
            documentModel.Save();

            // If the button is disabled in the configuration, dont show pane and disable button
            if (buttonTestResultReport.Checked && !documentModel.Configuration.ConfigurationTest.ConfigurationTestResult.Available)
            {
                buttonTestResultReport.Enabled = false;
                buttonTestResultReport.Checked = false;
            }

            // Set visibility of pane to checked state of button
            TestResultReport taskPane = null;
            if (buttonTestResultReport.Checked)
            {
                var testAdapter = SyncServiceFactory.CreateTfsTestAdapter(documentModel.TfsServer, documentModel.TfsProject, documentModel.Configuration);
                taskPane = new TestResultReport(documentModel.WordDocument as Document);
                var wordToTfsViewDispatcher = new ViewDispatcher(taskPane.Dispatcher);
                taskPane.Model = new TestResultReportModel(documentModel, wordToTfsViewDispatcher, testAdapter, this, new TestReportingProgressCancellationService(true));
            }
            SetHostPaneVisibility(documentModel.WordDocument as Document, buttonTestResultReport.Checked, taskPane, Resources.TestResultReportPanelCaption);
        }

        /// <summary>
        /// Called when the user clicks the button 'Insert Area/Iteration'.
        /// </summary>
        private void HandleInsertAreasClick(object sender, RibbonControlEventArgs e)
        {
            var documentModel = WordApplicationModel.GetModel(Globals.ThisAddIn.Application.ActiveDocument);
            if (documentModel == null)
            {
                return;
            }

            AreaIterationPathView paneContent = null;
            if (buttonAreaIterationPath.Checked)
            {
                var tfsService = SyncServiceFactory.CreateTfsService(documentModel.TfsServer, documentModel.TfsProject, documentModel.Configuration);
                paneContent = new AreaIterationPathView
                                  {
                                      Model = new AreaIterationPathViewModel(tfsService, documentModel)
                                  };
            }
            SetHostPaneVisibility(documentModel.WordDocument as Document, buttonAreaIterationPath.Checked, paneContent, Resources.AreaIterationPathPanelCaption);
        }

        /// <summary>
        /// Get host pane for current word document.
        /// </summary>
        /// <typeparam name="TControl">Host pane control type.</typeparam>
        /// <param name="wordDocument">Word document for which to return host pane.</param>
        /// <returns></returns>
        private static HostPane<TControl> GetHostPaneForDocument<TControl>(Document wordDocument) where TControl : UIElement
        {
            foreach (var taskPane in Globals.ThisAddIn.WordToTfsTaskPanes)
            {
                var hostPane = taskPane.Control as HostPane<TControl>;
                if (hostPane != null && hostPane.HostedControl != null)
                {
                    var hostPaneControl = (IHostPaneControl)hostPane.HostedControl;
                    if (hostPaneControl.AttachedDocument.ActiveWindow == wordDocument.ActiveWindow)
                    {
                        return hostPane;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Sets the visibility of the task pane for the current document
        /// </summary>
        /// <typeparam name="TControl">Type of the hosted pane.</typeparam>
        /// <param name="wordDocument">Word document for which to return host pane.</param>
        /// <param name="visibility">New visibility of the host pane</param>
        /// <param name="content">The content to show if visibility if set to true. Can be omitted if visibility is set to false.</param>
        /// <param name="title">The title of the host pane if visibility if set to true. Can be omitted if visibility is set to false.</param>
        private static void SetHostPaneVisibility<TControl>(Document wordDocument, bool visibility, TControl content = null, string title = null) where TControl : UIElement
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                var pane = GetHostPaneForDocument<TControl>(wordDocument);

                if (pane == null)
                {
                    if (visibility == false) return;
                    pane = new HostPane<TControl>(content);
                }
                else
                {
                    if (visibility || content != null)
                    {
                        pane.HostedControl = content;
                    }
                }

                if (visibility)
                {
                    Globals.ThisAddIn.ShowTaskPane(wordDocument, pane, title);
                }
                else
                {
                    Globals.ThisAddIn.HideTaskPane(wordDocument, pane);
                }
            }
            catch (COMException e)
            {
                SyncServiceTrace.LogException(e);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        #endregion

        #region Refresh process related methods

        /// <summary>
        /// Initiates the refresh process.
        /// </summary>
        /// <param name="documentModel">Model of associated word document.</param>
        /// <param name="workItems">Work items to refresh</param>
        private void InitiateRefresh(ISyncServiceDocumentModel documentModel, IEnumerable<IWorkItem> workItems)
        {
            documentModel.OperationInProgress = true;
            documentModel.Save();
            Cursor.Current = Cursors.WaitCursor;
            SyncServiceTrace.I("Starting refresh. Selected template = {0}", documentModel.MappingShowName);

            var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
            if (infoStorage != null)
            {
                infoStorage.ClearAll();
            }

            // Disable all other buttons during the operation
            DisableAllControls(documentModel);

            // display progress dialog
            var progressService = SyncServiceFactory.GetService<IProgressService>();

            if (!progressService.IsVisibleProgressWindow)
            {
                progressService.NewProgress(Resources.RefreshWorkItemsCaption);
                progressService.ShowProgress();

            }

            // start background thread to execute the import
            var backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += DoRefresh;
            backgroundWorker.RunWorkerCompleted += (s, e) =>
                                                       {
                                                           backgroundWorker.Dispose();
                                                           RefreshFinished(documentModel, e);
                                                       };
            backgroundWorker.RunWorkerAsync(new PublishArguments
                                                {
                                                    WorkItems = workItems,
                                                    DocumentModel = documentModel,
                                                });
        }

        /// <summary>
        /// Background method to refresh the work items.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A <see cref="DoWorkEventArgs"/> that contains the event data. If you want only
        /// the selected item to be refreshed, pass True as parameter to the background worker</param>
        private void DoRefresh(object sender, DoWorkEventArgs e)
        {
            SyncServiceTrace.I("Refreshing work items...");

            // intialize services
            var workItemSyncService = SyncServiceFactory.GetService<IWorkItemSyncService>();
            var arguments = (PublishArguments)e.Argument;

            // set actual selected configuration
            arguments.DocumentModel.Configuration.ActivateMapping(arguments.DocumentModel.MappingShowName);
            if (arguments.DocumentModel.Configuration.EnableSwitchFormatting)
            {
                arguments.DocumentModel.Configuration.IgnoreFormatting = chkIgnoreFormatting.Checked;
            }
            if (arguments.DocumentModel.Configuration.EnableSwitchConflictOverwrite)
            {
                arguments.DocumentModel.Configuration.ConflictOverwrite = chkConflictOverwrite.Checked;
            }

            // get adapter
            var source = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(arguments.DocumentModel.TfsServer, arguments.DocumentModel.TfsProject, null, arguments.DocumentModel.Configuration);
            _refreshWordAdapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(arguments.DocumentModel.WordDocument, arguments.DocumentModel.Configuration);
            _refreshWordAdapter.PrepareDocumentForLongTermOperation();
            var selectedItems = arguments.WorkItems;

            var destinationIds = (from IWorkItem item in selectedItems
                                  where item.Id > 0
                                  select item.Id).ToArray();

            // Get the work items that are actually mapped to TFS
            IList<IWorkItem> sourceWorkItems;
            if (ValidateRefresh(destinationIds, source, out sourceWorkItems) == false)
            {
                return;
            }

            // synchronize Word document work items from/with TFS work items
            workItemSyncService.Refresh(source, _refreshWordAdapter, sourceWorkItems, arguments.DocumentModel.Configuration);
        }

        /// <summary>
        /// Occurs when the background operation has completed, has been canceled, or has raised an exception.
        /// </summary>
        private void RefreshFinished(ISyncServiceDocumentModel documentModel, RunWorkerCompletedEventArgs e)
        {
            e.Error.NotifyIfException("Failed to refresh work items.");

            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.HideProgress();
            Cursor.Current = Cursors.Default;

            _refreshWordAdapter.UndoPreparationsDocumentForLongTermOperation();

            documentModel.OperationInProgress = false;
            EnableAllControls(documentModel);
            SyncServiceTrace.I("Refreshing work items finished.");
        }

        #endregion Refresh process related methods

        #region Publish process related methods

        private void InitiatePublish(ISyncServiceDocumentModel documentModel, IEnumerable<IWorkItem> workItems)
        {
            documentModel.OperationInProgress = true;

            Cursor.Current = Cursors.WaitCursor;
            documentModel.Configuration.ActivateMapping(documentModel.MappingShowName);
            documentModel.Save();
            SyncServiceTrace.I("Starting publish. Selected template = {0}", documentModel.MappingShowName);

            if (documentModel.Configuration.EnableSwitchFormatting)
            {
                documentModel.Configuration.IgnoreFormatting = chkIgnoreFormatting.Checked;
            }
            if (documentModel.Configuration.EnableSwitchConflictOverwrite)
            {
                documentModel.Configuration.ConflictOverwrite = chkConflictOverwrite.Checked;
            }

            // that one is passed as parameter
            //if (configService.EnableSwitchConflictOverwriting) configService.ForceConflictOverwrite = chkForcePublish.Checked;
            var forcePublish = documentModel.Configuration.EnableSwitchConflictOverwrite ? chkConflictOverwrite.Checked : documentModel.Configuration.ConflictOverwrite;

            var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
            if (infoStorage != null)
            {
                infoStorage.ClearAll();
            }

            // Disable all other buttons during the operation
            DisableAllControls(documentModel);

            // Show the publishing pane
            // TODO Using a module field here is wrong. Different documents should be bound to different instances of the publishing pane -SER
            var document = documentModel.WordDocument as Document;
            _publishingStateControl = new PublishingStateControl(document)
                                          {
                                              IsInProgress = true
                                          };

            SetHostPaneVisibility(document, true, _publishingStateControl, Resources.PublishTitle);

            // display progress dialog
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.ShowProgress();
            progressService.NewProgress(Resources.LO_Publish_Title);

            //Use the Fieldverifier to check the types of the work items thar are about to get published.
            //Compares the fields with the corresponding fields on the server. This is necessary to avoid formatting errors and wrong published fields
            var veryfier = new FieldVerifier();
            var destination = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(documentModel.TfsServer, documentModel.TfsProject, null, documentModel.Configuration);
            if (veryfier.VerifyTemplateMapping(workItems, destination, documentModel.Configuration))
            {
                var overwriteWindow = new OverwriteDialog();
                overwriteWindow.ShowDialog();
                var result = overwriteWindow.DialogResult;
                if (result == true)
                {
                    //Users Decision is "true", set the flag for each field to trigger an overwrite
                    foreach (var workItemDictionary in veryfier.WrongMappedFieldsDictionary)
                    {
                        //Loop through all WorkItems that should be published
                        foreach (var singleWorkItem in workItems)
                        {
                            //Look if it is in the dictionary
                            if (singleWorkItem == workItemDictionary.Key)
                            {
                                //Adjust all fields
                                foreach (var field in workItemDictionary.Value)
                                {
                                    if (field.GetType() == typeof(WordTableField))
                                    {
                                        if (singleWorkItem.Fields.Contains(field))
                                        {

                                            //WordTableField thisField = singleWorkItem.Fields[field.ReferenceName] as WordTableField;
                                            //thisField.ParseHtmlAsPlaintext = field.ParseHtmlAsPlaintext;

                                            ((WordTableField)(singleWorkItem.Fields[field.ReferenceName])).ParseHtmlAsPlaintext = field.ParseHtmlAsPlaintext;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            // start background thread
            var bw = new BackgroundWorker
                         {
                             WorkerSupportsCancellation = true
                         };
            bw.DoWork += DoPublish;

            // I don't know why I do have to manually invoke with the dispatcher, but the
            // AsyncOperationManager.SynchronizationContext is lost after the first user event,
            // causing the main thread to be abandoned and thus all access to controls become invalid
            // because they where created in the abandoned Thread.
            bw.RunWorkerCompleted += (s, e) =>
                                         {
                                             bw.Dispose();
                                             _dispatcher.BeginInvoke((Action)(() => PublishFinished(documentModel, e)));
                                         };

            bw.RunWorkerAsync(new PublishArguments
                                  {
                                      WorkItems = workItems,
                                      ForcePublish = forcePublish,
                                      DocumentModel = documentModel,
                                  });
        }

        /// <summary>
        /// Background method to publish the work items. Occurs when RunWorkerAsync is called.
        /// </summary>
        private void DoPublish(object sender, DoWorkEventArgs e)
        {
            SyncServiceTrace.I("Publishing work items...");

            var arguments = (PublishArguments)e.Argument;

            var document = arguments.DocumentModel.WordDocument as Document;
            if (document == null)
            {
                return;
            }

            var trackChanges = document.TrackRevisions;
            if (trackChanges)
            {
                document.TrackRevisions = false;
            }

            var lastValueRelyOnVml = document.WebOptions.RelyOnVML;
            var lastValueShowRevisions = document.ShowRevisions;
            document.WebOptions.RelyOnVML = false;
            document.ShowRevisions = false;
            document.WebOptions.AllowPNG = true;

            try
            {
                var service = SyncServiceFactory.GetService<IWorkItemSyncService>();
                _publishWordAdapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(document, arguments.DocumentModel.Configuration);
                var destination = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(arguments.DocumentModel.TfsServer, arguments.DocumentModel.TfsProject, null, arguments.DocumentModel.Configuration);

                _publishWordAdapter.PrepareDocumentForLongTermOperation();
                service.Publish(_publishWordAdapter, destination, arguments.WorkItems, arguments.ForcePublish, arguments.DocumentModel.Configuration);
            }
            finally
            {
                arguments.DocumentModel.OperationInProgress = false;
                document.WebOptions.RelyOnVML = lastValueRelyOnVml;
                document.ShowRevisions = lastValueShowRevisions;
                document.TrackRevisions = trackChanges;
            }
        }

        /// <summary>
        /// Occurs when the background operation has completed, has been canceled, or has raised an exception.
        /// </summary>
        private void PublishFinished(ISyncServiceDocumentModel model, RunWorkerCompletedEventArgs e)
        {
            e.Error.NotifyIfException(Resources.Error_Publish);

            // hide progress window
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.HideProgress();

            _publishWordAdapter.UndoPreparationsDocumentForLongTermOperation();

            Cursor.Current = Cursors.Default;
            _publishingStateControl.IsInProgress = false;

            EnableAllControls(model);
            SyncServiceTrace.I("Publishing work items finished.");
        }

        #endregion Publish process related methods

        #region Implementation of IWordRibbon

        /// <summary>
        /// The method sets / resets all required members to appropriate state.
        /// </summary>
        /// <param name="documentModel">Model of the document where will be the operation started.</param>
        public void ResetBeforeOperation(ISyncServiceDocumentModel documentModel)
        {
        }

        /// <summary>
        /// Method sets the state of all controls to desired state that depends on the model - active document.
        /// </summary>
        public void EnableAllControls(ISyncServiceDocumentModel documentModel)
        {
            if (documentModel == null)
            {
                Debug.Assert(documentModel == null, "Document model not found!");
                return;
            }

            if (documentModel.OperationInProgress)
            {
                // Disable all other buttons during the operation
                DisableAllControls(documentModel);
                return;
            }

            // TFS group
            buttonBindUnbind.Enabled = true;
            labelServer.Enabled = true;
            labelServerText.Enabled = true;
            labelProject.Enabled = true;
            labelProjectText.Enabled = true;

            // Publish group
            splitButtonPublish.Enabled = documentModel.TfsDocumentBound && !documentModel.TestReportGenerated;
            splitButtonRefresh.Enabled = documentModel.TfsDocumentBound && !documentModel.TestReportGenerated;
            buttonGetWorkItems.Enabled = documentModel.TfsDocumentBound && !documentModel.TestReportGenerated;
            buttonGetWorkItems.Checked = Globals.ThisAddIn.TaskPaneVisible(documentModel.WordDocument as Document, Resources.GetWorkItemsPanelCaption);
            buttonSynchronizationState.Enabled = documentModel.TfsDocumentBound && !documentModel.TestReportGenerated;
            buttonSynchronizationState.Checked = Globals.ThisAddIn.TaskPaneVisible(documentModel.WordDocument as Document, Resources.SynchronizationStatePanelCaption);

            // Template group
            chkIgnoreFormatting.Enabled = true;
            chkConflictOverwrite.Enabled = true;
            buttonTemplateManager.Enabled = true;
            buttonTemplateManager.Checked = Globals.ThisAddIn.TaskPaneVisible(documentModel.WordDocument as Document, Resources.TemplateManagerPanelCaption);
            buttonInsertTemplate.Enabled = false;
            buttonCustomizeStyle.Enabled = false;
            buttonDeleteStyle.Enabled = false;

            // Work items group
            menuNewEmptyWorkItem.Enabled = true;
            menuNewEmptyWorkItem.ShowImage = true;
            menuNewWorkItem.Enabled = true;
            menuNewWorkItem.ShowImage = true;
            menuNewHeader.Enabled = true;
            menuNewHeader.ShowImage = true;
            buttonEditDefVals.Enabled = true;
            buttonEditDefVals.Checked = Globals.ThisAddIn.TaskPaneVisible(documentModel.WordDocument as Document, Resources.DefaultValuesPanelCaption);
            buttonDeleteIds.Enabled = true;
            buttonAreaIterationPath.Enabled = documentModel.TfsDocumentBound;
            buttonAreaIterationPath.Checked = Globals.ThisAddIn.TaskPaneVisible(documentModel.WordDocument as Document, Resources.AreaIterationPathPanelCaption);

            // Configuration group
            buttonPreserveReqID.Enabled = false;

            // Test management group
            splitButtonTestSpecificationReport.Enabled = documentModel.TfsDocumentBound;
            splitButtonTestSpecificationReport.Checked = Globals.ThisAddIn.TaskPaneVisible(documentModel.WordDocument as Document, Resources.TestSpecificationReportPanelCaption);
            buttonTestSpecificationReportByQuery.Enabled = documentModel.TfsDocumentBound;
            buttonTestSpecificationReportByQuery.Checked = Globals.ThisAddIn.TaskPaneVisible(documentModel.WordDocument as Document, Resources.TestSpecificationReportByQueryPanelCaption);
            buttonTestResultReport.Enabled = documentModel.TfsDocumentBound;
            buttonTestResultReport.Checked = Globals.ThisAddIn.TaskPaneVisible(documentModel.WordDocument as Document, Resources.TestResultReportPanelCaption);

            // Help group
            buttonAbout.Enabled = true;
            buttonHelp.Enabled = true;

            buttonSettings.Enabled = true;

            // Texts in TFS group
            if (documentModel.TfsDocumentBound)
            {
                buttonBindUnbind.Label = Resources.TFSBindTeamProjectDisconnectLabel;
                buttonBindUnbind.ScreenTip = Resources.TFSUnbindTeamProjectScreenTip;
                buttonBindUnbind.SuperTip = Resources.TFSUnbindTeamProjectSuperTip;
                labelServerText.Label = documentModel.TfsServer;
                labelProjectText.Label = documentModel.TfsProject;
            }
            else
            {
                buttonBindUnbind.Label = Resources.TFSBindTeamProjectLabel;
                buttonBindUnbind.ScreenTip = Resources.TFSBindTeamProjectScreenTip;
                buttonBindUnbind.SuperTip = Resources.TFSBindTeamProjectSuperTip;
                labelServerText.Label = Resources.TFSServerDisconnected;
                labelProjectText.Label = Resources.TFSProjectDisconnected;
            }

            dropDownSelectTemplate.Enabled = (dropDownSelectTemplate.Items.Count > 0 && documentModel.Configuration.EnableTemplateSelection);
        }

        /// <summary>
        /// Disables all controls. Used if active document is not valid.
        /// </summary>
        public void DisableAllControls(ISyncServiceDocumentModel documentModel)
        {
            foreach (var group in tabAITRM.Groups)
            {
                if (group == groupHelp)
                {
                    continue;
                }

                foreach (var control in group.Items)
                {
                    control.Enabled = false;
                }
            }

            // Texts in TFS group
            if (documentModel != null && documentModel.TfsDocumentBound)
            {
                buttonBindUnbind.Label = Resources.TFSUnbindTeamProjectLabel;
                buttonBindUnbind.ScreenTip = Resources.TFSUnbindTeamProjectScreenTip;
                buttonBindUnbind.SuperTip = Resources.TFSUnbindTeamProjectSuperTip;
                labelServerText.Label = documentModel.TfsServer;
                labelProjectText.Label = documentModel.TfsProject;
            }
            else
            {
                buttonBindUnbind.Label = Resources.TFSBindTeamProjectLabel;
                buttonBindUnbind.ScreenTip = Resources.TFSBindTeamProjectScreenTip;
                buttonBindUnbind.SuperTip = Resources.TFSBindTeamProjectSuperTip;
                labelServerText.Label = Resources.TFSServerDisconnected;
                labelProjectText.Label = Resources.TFSProjectDisconnected;
            }
        }

        #endregion Implementation of IWordRibbon

        #region Implementation of IWin32Window

        /// <summary>
        /// Gets the handle to the window represented by the implementer.
        /// </summary>
        public IntPtr Handle
        {
            get
            {
                return NativeMethods.FindWindow(WordClassName, Globals.ThisAddIn.Application.ActiveWindow.Caption + " - " + Globals.ThisAddIn.Application.Caption);
            }
        }

        #endregion Implementation of IWin32Window
    }
}