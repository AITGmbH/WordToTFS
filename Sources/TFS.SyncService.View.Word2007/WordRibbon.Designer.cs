namespace TFS.SyncService.View.Word2007
{
    partial class WordRibbon
    {
        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WordRibbon));
            this.tabAITRM = this.Factory.CreateRibbonTab();
            this.groupTFS = this.Factory.CreateRibbonGroup();
            this.buttonBindUnbind = this.Factory.CreateRibbonButton();
            this.boxServer = this.Factory.CreateRibbonBox();
            this.labelServer = this.Factory.CreateRibbonLabel();
            this.labelServerText = this.Factory.CreateRibbonLabel();
            this.boxProject = this.Factory.CreateRibbonBox();
            this.labelProject = this.Factory.CreateRibbonLabel();
            this.labelProjectText = this.Factory.CreateRibbonLabel();
            this.groupPublish = this.Factory.CreateRibbonGroup();
            this.splitButtonPublish = this.Factory.CreateRibbonSplitButton();
            this.buttonPublishSelected = this.Factory.CreateRibbonButton();
            this.buttonForcePublish = this.Factory.CreateRibbonButton();
            this.splitButtonRefresh = this.Factory.CreateRibbonSplitButton();
            this.buttonRefreshSelected = this.Factory.CreateRibbonButton();
            this.buttonGetWorkItems = this.Factory.CreateRibbonToggleButton();
            this.buttonSynchronizationState = this.Factory.CreateRibbonToggleButton();
            this.groupTemplate = this.Factory.CreateRibbonGroup();
            this.dropDownSelectTemplate = this.Factory.CreateRibbonDropDown();
            this.chkIgnoreFormatting = this.Factory.CreateRibbonCheckBox();
            this.chkConflictOverwrite = this.Factory.CreateRibbonCheckBox();
            this.buttonTemplateManager = this.Factory.CreateRibbonToggleButton();
            this.buttonInsertTemplate = this.Factory.CreateRibbonButton();
            this.buttonCustomizeStyle = this.Factory.CreateRibbonButton();
            this.buttonDeleteStyle = this.Factory.CreateRibbonButton();
            this.groupRequirements = this.Factory.CreateRibbonGroup();
            this.menuNewEmptyWorkItem = this.Factory.CreateRibbonMenu();
            this.menuNewWorkItem = this.Factory.CreateRibbonMenu();
            this.menuNewHeader = this.Factory.CreateRibbonMenu();
            this.buttonEditDefVals = this.Factory.CreateRibbonToggleButton();
            this.buttonDeleteIds = this.Factory.CreateRibbonButton();
            this.buttonAreaIterationPath = this.Factory.CreateRibbonToggleButton();
            this.groupConfiguration = this.Factory.CreateRibbonGroup();
            this.buttonPreserveReqID = this.Factory.CreateRibbonToggleButton();
            this.groupTestManagement = this.Factory.CreateRibbonGroup();
            this.splitButtonTestSpecificationReport = this.Factory.CreateRibbonSplitButton();
            this.buttonTestSpecificationReportByQuery = this.Factory.CreateRibbonToggleButton();
            this.buttonTestResultReport = this.Factory.CreateRibbonToggleButton();
            this.groupHelp = this.Factory.CreateRibbonGroup();
            this.buttonHelp = this.Factory.CreateRibbonButton();
            this.buttonAbout = this.Factory.CreateRibbonButton();
            this.buttonUpdate = this.Factory.CreateRibbonButton();
            this.FeedbackButton = this.Factory.CreateRibbonButton();
            this.groupSettings = this.Factory.CreateRibbonGroup();
            this.buttonSettings = this.Factory.CreateRibbonButton();
            this.tabAITRM.SuspendLayout();
            this.groupTFS.SuspendLayout();
            this.boxServer.SuspendLayout();
            this.boxProject.SuspendLayout();
            this.groupPublish.SuspendLayout();
            this.groupTemplate.SuspendLayout();
            this.groupRequirements.SuspendLayout();
            this.groupConfiguration.SuspendLayout();
            this.groupTestManagement.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.groupSettings.SuspendLayout();
            // 
            // tabAITRM
            // 
            this.tabAITRM.Groups.Add(this.groupTFS);
            this.tabAITRM.Groups.Add(this.groupPublish);
            this.tabAITRM.Groups.Add(this.groupTemplate);
            this.tabAITRM.Groups.Add(this.groupRequirements);
            this.tabAITRM.Groups.Add(this.groupConfiguration);
            this.tabAITRM.Groups.Add(this.groupTestManagement);
            this.tabAITRM.Groups.Add(this.groupHelp);
            this.tabAITRM.Groups.Add(this.groupSettings);
            resources.ApplyResources(this.tabAITRM, "tabAITRM");
            this.tabAITRM.Name = "tabAITRM";
            // 
            // groupTFS
            // 
            this.groupTFS.Items.Add(this.buttonBindUnbind);
            this.groupTFS.Items.Add(this.boxServer);
            this.groupTFS.Items.Add(this.boxProject);
            resources.ApplyResources(this.groupTFS, "groupTFS");
            this.groupTFS.Name = "groupTFS";
            // 
            // buttonBindUnbind
            // 
            resources.ApplyResources(this.buttonBindUnbind, "buttonBindUnbind");
            this.buttonBindUnbind.Name = "buttonBindUnbind";
            this.buttonBindUnbind.ShowImage = true;
            this.buttonBindUnbind.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonBindUnbindClick);
            // 
            // boxServer
            // 
            this.boxServer.Items.Add(this.labelServer);
            this.boxServer.Items.Add(this.labelServerText);
            this.boxServer.Name = "boxServer";
            // 
            // labelServer
            // 
            resources.ApplyResources(this.labelServer, "labelServer");
            this.labelServer.Name = "labelServer";
            // 
            // labelServerText
            // 
            resources.ApplyResources(this.labelServerText, "labelServerText");
            this.labelServerText.Name = "labelServerText";
            // 
            // boxProject
            // 
            this.boxProject.Items.Add(this.labelProject);
            this.boxProject.Items.Add(this.labelProjectText);
            this.boxProject.Name = "boxProject";
            // 
            // labelProject
            // 
            resources.ApplyResources(this.labelProject, "labelProject");
            this.labelProject.Name = "labelProject";
            // 
            // labelProjectText
            // 
            resources.ApplyResources(this.labelProjectText, "labelProjectText");
            this.labelProjectText.Name = "labelProjectText";
            // 
            // groupPublish
            // 
            this.groupPublish.Items.Add(this.splitButtonPublish);
            this.groupPublish.Items.Add(this.splitButtonRefresh);
            this.groupPublish.Items.Add(this.buttonGetWorkItems);
            this.groupPublish.Items.Add(this.buttonSynchronizationState);
            resources.ApplyResources(this.groupPublish, "groupPublish");
            this.groupPublish.Name = "groupPublish";
            // 
            // splitButtonPublish
            // 
            this.splitButtonPublish.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButtonPublish.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.Publish;
            this.splitButtonPublish.Items.Add(this.buttonPublishSelected);
            this.splitButtonPublish.Items.Add(this.buttonForcePublish);
            resources.ApplyResources(this.splitButtonPublish, "splitButtonPublish");
            this.splitButtonPublish.Name = "splitButtonPublish";
            this.splitButtonPublish.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonPublishClick);
            // 
            // buttonPublishSelected
            // 
            this.buttonPublishSelected.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.Publish;
            resources.ApplyResources(this.buttonPublishSelected, "buttonPublishSelected");
            this.buttonPublishSelected.Name = "buttonPublishSelected";
            this.buttonPublishSelected.ShowImage = true;
            this.buttonPublishSelected.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonPublishSelectedClick);
            // 
            // buttonForcePublish
            // 
            this.buttonForcePublish.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.Publish;
            resources.ApplyResources(this.buttonForcePublish, "buttonForcePublish");
            this.buttonForcePublish.Name = "buttonForcePublish";
            this.buttonForcePublish.ShowImage = true;
            // 
            // splitButtonRefresh
            // 
            this.splitButtonRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButtonRefresh.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.table_refresh;
            this.splitButtonRefresh.Items.Add(this.buttonRefreshSelected);
            resources.ApplyResources(this.splitButtonRefresh, "splitButtonRefresh");
            this.splitButtonRefresh.Name = "splitButtonRefresh";
            this.splitButtonRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRefreshClick);
            // 
            // buttonRefreshSelected
            // 
            this.buttonRefreshSelected.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.table_refresh;
            resources.ApplyResources(this.buttonRefreshSelected, "buttonRefreshSelected");
            this.buttonRefreshSelected.Name = "buttonRefreshSelected";
            this.buttonRefreshSelected.ShowImage = true;
            this.buttonRefreshSelected.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRefreshSelectedClick);
            // 
            // buttonGetWorkItems
            // 
            this.buttonGetWorkItems.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonGetWorkItems.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.GetWorkItems;
            resources.ApplyResources(this.buttonGetWorkItems, "buttonGetWorkItems");
            this.buttonGetWorkItems.Name = "buttonGetWorkItems";
            this.buttonGetWorkItems.ShowImage = true;
            this.buttonGetWorkItems.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonGetWorkItemsClick);
            // 
            // buttonSynchronizationState
            // 
            this.buttonSynchronizationState.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSynchronizationState.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.Preserve;
            resources.ApplyResources(this.buttonSynchronizationState, "buttonSynchronizationState");
            this.buttonSynchronizationState.Name = "buttonSynchronizationState";
            this.buttonSynchronizationState.ShowImage = true;
            this.buttonSynchronizationState.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonSynchronizationStateClick);
            // 
            // groupTemplate
            // 
            this.groupTemplate.Items.Add(this.dropDownSelectTemplate);
            this.groupTemplate.Items.Add(this.chkIgnoreFormatting);
            this.groupTemplate.Items.Add(this.chkConflictOverwrite);
            this.groupTemplate.Items.Add(this.buttonTemplateManager);
            this.groupTemplate.Items.Add(this.buttonInsertTemplate);
            this.groupTemplate.Items.Add(this.buttonCustomizeStyle);
            this.groupTemplate.Items.Add(this.buttonDeleteStyle);
            resources.ApplyResources(this.groupTemplate, "groupTemplate");
            this.groupTemplate.Name = "groupTemplate";
            // 
            // dropDownSelectTemplate
            // 
            resources.ApplyResources(this.dropDownSelectTemplate, "dropDownSelectTemplate");
            this.dropDownSelectTemplate.Name = "dropDownSelectTemplate";
            this.dropDownSelectTemplate.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleDropDownSelectTemplateSelectionChanged);
            // 
            // chkIgnoreFormatting
            // 
            resources.ApplyResources(this.chkIgnoreFormatting, "chkIgnoreFormatting");
            this.chkIgnoreFormatting.Name = "chkIgnoreFormatting";
            this.chkIgnoreFormatting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleCheckboxIgnoreFormattingClick);
            // 
            // chkConflictOverwrite
            // 
            resources.ApplyResources(this.chkConflictOverwrite, "chkConflictOverwrite");
            this.chkConflictOverwrite.Name = "chkConflictOverwrite";
            // 
            // buttonTemplateManager
            // 
            this.buttonTemplateManager.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonTemplateManager.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.StyleCustomize;
            resources.ApplyResources(this.buttonTemplateManager, "buttonTemplateManager");
            this.buttonTemplateManager.Name = "buttonTemplateManager";
            this.buttonTemplateManager.ShowImage = true;
            this.buttonTemplateManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonTemplateManagerClick);
            // 
            // buttonInsertTemplate
            // 
            this.buttonInsertTemplate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonInsertTemplate.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.InsertTemplate;
            resources.ApplyResources(this.buttonInsertTemplate, "buttonInsertTemplate");
            this.buttonInsertTemplate.Name = "buttonInsertTemplate";
            this.buttonInsertTemplate.ShowImage = true;
            // 
            // buttonCustomizeStyle
            // 
            this.buttonCustomizeStyle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonCustomizeStyle.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.StyleCustomize;
            resources.ApplyResources(this.buttonCustomizeStyle, "buttonCustomizeStyle");
            this.buttonCustomizeStyle.Name = "buttonCustomizeStyle";
            this.buttonCustomizeStyle.ShowImage = true;
            // 
            // buttonDeleteStyle
            // 
            this.buttonDeleteStyle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonDeleteStyle.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.StyleDelete;
            resources.ApplyResources(this.buttonDeleteStyle, "buttonDeleteStyle");
            this.buttonDeleteStyle.Name = "buttonDeleteStyle";
            this.buttonDeleteStyle.ShowImage = true;
            // 
            // groupRequirements
            // 
            this.groupRequirements.Items.Add(this.menuNewEmptyWorkItem);
            this.groupRequirements.Items.Add(this.menuNewWorkItem);
            this.groupRequirements.Items.Add(this.menuNewHeader);
            this.groupRequirements.Items.Add(this.buttonEditDefVals);
            this.groupRequirements.Items.Add(this.buttonDeleteIds);
            this.groupRequirements.Items.Add(this.buttonAreaIterationPath);
            resources.ApplyResources(this.groupRequirements, "groupRequirements");
            this.groupRequirements.Name = "groupRequirements";
            // 
            // menuNewEmptyWorkItem
            // 
            this.menuNewEmptyWorkItem.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuNewEmptyWorkItem.Dynamic = true;
            this.menuNewEmptyWorkItem.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.standard;
            resources.ApplyResources(this.menuNewEmptyWorkItem, "menuNewEmptyWorkItem");
            this.menuNewEmptyWorkItem.Name = "menuNewEmptyWorkItem";
            this.menuNewEmptyWorkItem.ShowImage = true;
            this.menuNewEmptyWorkItem.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleMenuNewEmptyWorkItemClick);
            // 
            // menuNewWorkItem
            // 
            this.menuNewWorkItem.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuNewWorkItem.Dynamic = true;
            this.menuNewWorkItem.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.standard;
            resources.ApplyResources(this.menuNewWorkItem, "menuNewWorkItem");
            this.menuNewWorkItem.Name = "menuNewWorkItem";
            this.menuNewWorkItem.ShowImage = true;
            this.menuNewWorkItem.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleMenuNewWorkItemClick);
            // 
            // menuNewHeader
            // 
            this.menuNewHeader.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuNewHeader.Dynamic = true;
            this.menuNewHeader.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.standard;
            resources.ApplyResources(this.menuNewHeader, "menuNewHeader");
            this.menuNewHeader.Name = "menuNewHeader";
            this.menuNewHeader.ShowImage = true;
            // 
            // buttonEditDefVals
            // 
            this.buttonEditDefVals.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonEditDefVals.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.ReqDefVals;
            resources.ApplyResources(this.buttonEditDefVals, "buttonEditDefVals");
            this.buttonEditDefVals.Name = "buttonEditDefVals";
            this.buttonEditDefVals.ShowImage = true;
            this.buttonEditDefVals.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonEditDefValsClick);
            // 
            // buttonDeleteIds
            // 
            this.buttonDeleteIds.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonDeleteIds.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.IdDelete;
            resources.ApplyResources(this.buttonDeleteIds, "buttonDeleteIds");
            this.buttonDeleteIds.Name = "buttonDeleteIds";
            this.buttonDeleteIds.ShowImage = true;
            this.buttonDeleteIds.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonDeleteIdsClick);
            // 
            // buttonAreaIterationPath
            // 
            this.buttonAreaIterationPath.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAreaIterationPath.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.branch;
            resources.ApplyResources(this.buttonAreaIterationPath, "buttonAreaIterationPath");
            this.buttonAreaIterationPath.Name = "buttonAreaIterationPath";
            this.buttonAreaIterationPath.ShowImage = true;
            this.buttonAreaIterationPath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleInsertAreasClick);
            // 
            // groupConfiguration
            // 
            this.groupConfiguration.Items.Add(this.buttonPreserveReqID);
            resources.ApplyResources(this.groupConfiguration, "groupConfiguration");
            this.groupConfiguration.Name = "groupConfiguration";
            // 
            // buttonPreserveReqID
            // 
            this.buttonPreserveReqID.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonPreserveReqID.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.Preserve;
            resources.ApplyResources(this.buttonPreserveReqID, "buttonPreserveReqID");
            this.buttonPreserveReqID.Name = "buttonPreserveReqID";
            this.buttonPreserveReqID.ShowImage = true;
            // 
            // groupTestManagement
            // 
            this.groupTestManagement.Items.Add(this.splitButtonTestSpecificationReport);
            this.groupTestManagement.Items.Add(this.buttonTestResultReport);
            resources.ApplyResources(this.groupTestManagement, "groupTestManagement");
            this.groupTestManagement.Name = "groupTestManagement";
            // 
            // splitButtonTestSpecificationReport
            // 
            this.splitButtonTestSpecificationReport.ButtonType = Microsoft.Office.Tools.Ribbon.RibbonButtonType.ToggleButton;
            this.splitButtonTestSpecificationReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButtonTestSpecificationReport.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.TestSpecificationReport;
            this.splitButtonTestSpecificationReport.Items.Add(this.buttonTestSpecificationReportByQuery);
            resources.ApplyResources(this.splitButtonTestSpecificationReport, "splitButtonTestSpecificationReport");
            this.splitButtonTestSpecificationReport.Name = "splitButtonTestSpecificationReport";
            this.splitButtonTestSpecificationReport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonTestSpecificationReportClick);
            // 
            // buttonTestSpecificationReportByQuery
            // 
            this.buttonTestSpecificationReportByQuery.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.TestSpecificationReport;
            resources.ApplyResources(this.buttonTestSpecificationReportByQuery, "buttonTestSpecificationReportByQuery");
            this.buttonTestSpecificationReportByQuery.Name = "buttonTestSpecificationReportByQuery";
            this.buttonTestSpecificationReportByQuery.ShowImage = true;
            this.buttonTestSpecificationReportByQuery.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonTestSpecificationReportByQueryClick);
            // 
            // buttonTestResultReport
            // 
            this.buttonTestResultReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonTestResultReport.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.TestResultReport;
            resources.ApplyResources(this.buttonTestResultReport, "buttonTestResultReport");
            this.buttonTestResultReport.Name = "buttonTestResultReport";
            this.buttonTestResultReport.ShowImage = true;
            this.buttonTestResultReport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonTestResultReportClick);
            // 
            // groupHelp
            // 
            this.groupHelp.Items.Add(this.buttonHelp);
            this.groupHelp.Items.Add(this.buttonAbout);
            this.groupHelp.Items.Add(this.buttonUpdate);
            this.groupHelp.Items.Add(this.FeedbackButton);
            resources.ApplyResources(this.groupHelp, "groupHelp");
            this.groupHelp.Name = "groupHelp";
            // 
            // buttonHelp
            // 
            this.buttonHelp.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.Help;
            resources.ApplyResources(this.buttonHelp, "buttonHelp");
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.ShowImage = true;
            this.buttonHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonHelpClick);
            // 
            // buttonAbout
            // 
            this.buttonAbout.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.About;
            resources.ApplyResources(this.buttonAbout, "buttonAbout");
            this.buttonAbout.Name = "buttonAbout";
            this.buttonAbout.ShowImage = true;
            this.buttonAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonAboutClick);
            // 
            // buttonUpdate
            // 
            this.buttonUpdate.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.BindUnbind;
            resources.ApplyResources(this.buttonUpdate, "buttonUpdate");
            this.buttonUpdate.Name = "buttonUpdate";
            this.buttonUpdate.ShowImage = true;
            this.buttonUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonUpdateClick);
            // 
            // FeedbackButton
            // 
            this.FeedbackButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FeedbackButton.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.aitIcon;
            resources.ApplyResources(this.FeedbackButton, "FeedbackButton");
            this.FeedbackButton.Name = "FeedbackButton";
            this.FeedbackButton.ShowImage = true;
            this.FeedbackButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleFeedbackButtonClick);

            //
            //Group Settings
            ///
            this.groupSettings.Items.Add(this.buttonSettings);

            this.buttonSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSettings.Image = global::TFS.SyncService.View.Word2007.Properties.Resources.Settings;
            resources.ApplyResources(this.buttonSettings, "ButtonSettings");
            this.buttonSettings.Name = "ButtonSettings";
            this.buttonSettings.ShowImage = true;
            this.buttonSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleButtonSettingsClick);
 

            // 
            // WordRibbon
            // 
            this.Name = "WordRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabAITRM);
            this.tabAITRM.ResumeLayout(false);
            this.tabAITRM.PerformLayout();
            this.groupTFS.ResumeLayout(false);
            this.groupTFS.PerformLayout();
            this.boxServer.ResumeLayout(false);
            this.boxServer.PerformLayout();
            this.boxProject.ResumeLayout(false);
            this.boxProject.PerformLayout();
            this.groupPublish.ResumeLayout(false);
            this.groupPublish.PerformLayout();
            this.groupTemplate.ResumeLayout(false);
            this.groupTemplate.PerformLayout();
            this.groupRequirements.ResumeLayout(false);
            this.groupRequirements.PerformLayout();
            this.groupConfiguration.ResumeLayout(false);
            this.groupConfiguration.PerformLayout();
            this.groupTestManagement.ResumeLayout(false);
            this.groupTestManagement.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabAITRM;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTFS;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPublish;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupRequirements;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupConfiguration;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTestManagement;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonTestSpecificationReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton buttonTestSpecificationReportByQuery;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton buttonTestResultReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton buttonPreserveReqID;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton buttonEditDefVals;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonInsertTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCustomizeStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeleteStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonBindUnbind;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxServer;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxProject;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelServer;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelServerText;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelProject;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelProjectText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeleteIds;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownSelectTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton buttonTemplateManager;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FeedbackButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuNewEmptyWorkItem;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuNewWorkItem;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton buttonGetWorkItems;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton buttonAreaIterationPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonRefresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonPublish;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonForcePublish;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPublishSelected;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkIgnoreFormatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkConflictOverwrite;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRefreshSelected;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuNewHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton buttonSynchronizationState;

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSettings;


    }
}
