

namespace TFS.SyncService.Test.Unit
{
    #region Usings
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.Linq;
    using AIT.TFS.SyncService.Adapter.TFS2012;
    using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
    using AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects;
    using AIT.TFS.SyncService.Contracts;
    using AIT.TFS.SyncService.Contracts.Adapter;
    using AIT.TFS.SyncService.Contracts.Configuration;
    using AIT.TFS.SyncService.Contracts.Enums;
    using AIT.TFS.SyncService.Contracts.InfoStorage;
    using AIT.TFS.SyncService.Contracts.TestCenter;
    using AIT.TFS.SyncService.Contracts.WorkItemCollections;
    using AIT.TFS.SyncService.Contracts.WorkItemObjects;
    using AIT.TFS.SyncService.Factory;
    using AIT.TFS.SyncService.Service;
    using AIT.TFS.SyncService.Service.Configuration;
    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.TestManagement.Client;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using TFS.Test.Common;
    using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
    using Microsoft.Office.Interop.Word;
    using System.IO;
    using System.Reflection;
    #endregion

    /// <summary>
    ///This is a test class for WorkItemSyncServiceTest and is intended
    ///to contain all WorkItemSyncServiceTest Unit Tests
    ///</summary>
    [TestClass]
    [DeploymentItem("Configuration", "Configuration")]
    public class WorkItemSyncServiceTest
    {
        #region Test initializations

        private Mock<IWordSyncAdapter> _wordAdapter;
        private Mock<IWorkItemSyncAdapter> _tfsAdapter;
        private Mock<ITfsService> _tfsServiceAdapter;
        private WorkItemSyncService _workItemSyncService;
        private static TestContext _testContext;

        #region Properties
        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext { get; set; }
        #endregion

        #region Test initializations

        /// <summary>
        /// Make sure all services and adapter creators are registered.
        /// </summary>
        [ClassInitialize]
        public static void MyClassInitialize(TestContext testContext)
        {
            _testContext = testContext;
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);

            AIT.TFS.SyncService.Adapter.TFS2012.AssemblyInit.Instance.Init();
            AIT.TFS.SyncService.Service.AssemblyInit.Instance.Init();

            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));
        }

        /// <summary>
        /// Create mock adapters and the SUT
        /// </summary>
        [TestInitialize]
        public void MyTestInitialize()
        {
            var systemVariableMock = new Mock<ISystemVariableService>();
            _wordAdapter = new Mock<IWordSyncAdapter> { DefaultValue = DefaultValue.Mock };
            _wordAdapter.Setup(x => x.Open(It.IsAny<int[]>())).Returns(true);
            _wordAdapter.SetupGet(x => x.WorkItems).Returns(new WorkItemCollection());

            _tfsAdapter = new Mock<IWorkItemSyncAdapter> { DefaultValue = DefaultValue.Mock };
            _tfsServiceAdapter = _tfsAdapter.As<ITfsService>();
            _tfsServiceAdapter.SetupGet(x => x.ProjectName).Returns("UnitTestProjectMock");
            _tfsServiceAdapter.Setup(x => x.GetRevisionWebAccessUri(It.IsAny<IWorkItem>(), It.IsAny<int>())).Returns(new Uri("http://www.google.de"));
            _tfsAdapter.Setup(x => x.Open(It.IsAny<int[]>())).Returns(true);
            _tfsAdapter.Setup(x => x.OpenWithConfigurations(It.IsAny<Dictionary<int, IConfigurationItem>>())).Returns(true);
            _tfsAdapter.SetupGet(x => x.WorkItems).Returns(new WorkItemCollection());
            _tfsAdapter.Setup(x => x.GetWorkItemEditorUrl(It.IsAny<int>())).Returns(new Uri("http://www.google.de"));

            _workItemSyncService = new WorkItemSyncService(systemVariableMock.Object);

            var informationStorage = SyncServiceFactory.GetService<IInfoStorageService>();
            informationStorage.UserInformation.Clear();
        }

        #endregion Test initializations

        #region Tests

        /// <summary>
        /// Tests whether test steps are transformed from and to a string representation of the format "Action #StepDelimiter #Result"
        ///adding the newtonsoft json as deployment item to make sure the test runs
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        [DeploymentItem(@"Newtonsoft.Json.dll")]
        public void WorkItemSyncService_SynchronizeTestSteps()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = (TfsWorkItem)adapter.WorkItems.Find(CommonConfiguration.TestCaseId);

            var steps = testCase.Fields["Microsoft.VSTS.TCM.Steps"];
            var template = "{0}{1}{2}\n{3}{4}{5}";
            var g1 = Guid.NewGuid();
            var g2 = Guid.NewGuid();
            var g3 = Guid.NewGuid();
            var g4 = Guid.NewGuid();
            var value = string.Format(CultureInfo.InvariantCulture, template, g1, steps.Configuration.TestCaseStepDelimiter, g2, g3, steps.Configuration.TestCaseStepDelimiter, g4);

            steps.Value = value;

            Assert.AreEqual(0, adapter.Save().Count);
            adapter.Close();

            // Load again with changes
            adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            testCase = (TfsWorkItem)adapter.WorkItems.Find(CommonConfiguration.TestCaseId);

            // make sure test steps are saved in the test manager
            Assert.AreEqual("<DIV><P>" + g1.ToString() + "</P></DIV>", ((ITestStep)testCase.TestCase.Actions[0]).Title.ToString());
            Assert.AreEqual("<DIV><P>" + g2.ToString() + "</P></DIV>", ((ITestStep)testCase.TestCase.Actions[0]).ExpectedResult.ToString());
            Assert.AreEqual("<DIV><P>" + g3.ToString() + "</P></DIV>", ((ITestStep)testCase.TestCase.Actions[1]).Title.ToString());
            Assert.AreEqual("<DIV><P>" + g4.ToString() + "</P></DIV>", ((ITestStep)testCase.TestCase.Actions[1]).ExpectedResult.ToString());
            Assert.AreEqual(value, testCase.Fields["Microsoft.VSTS.TCM.Steps"].Value);
        }

        /// <summary>
        /// Tests whether AreaPaths have the TeamProjectName added when publishing
        ///</summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldAddProjectNameToAreaPath()
        {
            var configuration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.AreaPath");
            var configurationItem = configuration.GetConfigurationItems().Single();

            var workItemSource = CreateWorkItem(configurationItem, 1, "SourceValue_");
            var workItemTarget = CreateWorkItem(configurationItem, 1, "TargetValue_");
            workItemSource.Object.Fields["System.AreaPath"].Value = "\\UI";

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource.Object }, false, configuration);

            // verify
            Assert.AreEqual(_tfsServiceAdapter.Object.ProjectName + "\\UI", workItemTarget.Object.Fields["System.AreaPath"].Value);
        }

        /// <summary>
        /// Tests whether AreaPaths have the TeamProjectName removed when refreshing
        ///</summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_ShouldRemoveProjectNameFromAreaPath()
        {
            var configuration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.AreaPath");
            var configurationItem = configuration.GetConfigurationItems().Single();

            var workItemSource = CreateWorkItem(configurationItem, 1, "SourceValue_");
            var workItemTarget = CreateWorkItem(configurationItem, 1, "TargetValue_");
            workItemTarget.Object.Fields["System.AreaPath"].Value = _tfsServiceAdapter.Object.ProjectName + "\\UI";

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { workItemTarget.Object }, configuration);

            // verify
            Assert.AreEqual("\\UI", workItemSource.Object.Fields["System.AreaPath"].Value);
        }

        /// <summary>
        /// Tests whether context dependent configuration items like header items that are applied while opening a document are synchronized.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldSynchronizeHeaderFields()
        {
            // make sure the base configuration contains less items than the actual configuration of the item in the source adapter
            IConfigurationItem configuration = CommonConfiguration.Configuration.GetConfigurationItems().First(x => x.WorkItemType == "Requirement").Clone();
            configuration.FieldConfigurations.Add(new ConfigurationFieldItem("System.MyTest", string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, null, string.Empty, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            var workItemSource = CreateWorkItem(configuration, 1, "SourceValue_");
            var workItemTarget = CreateWorkItem(configuration, 1, "TargetValue_");

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource.Object }, false, CommonConfiguration.Configuration);

            // verify
            Assert.AreEqual(workItemSource.Object.Fields["System.MyTest"].Value, workItemTarget.Object.Fields["System.MyTest"].Value);
        }

        /// <summary>
        /// If stack rank is set, all published work items should have an increasing stack rank
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldSetStackRank()
        {
            // make sure the base configuration contains less items than the actual configuration of the item in the source adapter
            var configuration = CommonConfiguration.MockConfiguration;

            // user stack rank and add field that otherwise would be added in the config constructor
            configuration.SetupGet(x => x.UseStackRank).Returns(true);
            var configurationItem = configuration.Object.GetConfigurationItems().First(x => x.WorkItemType == "Requirement").Clone();
            configurationItem.FieldConfigurations.Add(new ConfigurationFieldItem("Microsoft.VSTS.Common.StackRank", string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 0, 0, string.Empty, false, HandleAsDocumentType.All, null, string.Empty, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            var workItemSource1 = CreateWorkItem(configurationItem, 1, "SourceValue1_");
            var workItemTarget1 = CreateWorkItem(configurationItem, 1, "TargetValue1_");
            var workItemSource2 = CreateWorkItem(configurationItem, 2, "SourceValue2_");
            var workItemTarget2 = CreateWorkItem(configurationItem, 2, "TargetValue2_");

            _wordAdapter.Object.WorkItems.Add(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget1.Object);
            _wordAdapter.Object.WorkItems.Add(workItemSource2.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget2.Object);

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object, workItemSource2.Object }, false, configuration.Object);

            // verify
            Assert.AreEqual(workItemTarget1.Object.Fields["Microsoft.VSTS.Common.StackRank"].Value, "1");
            Assert.AreEqual(workItemTarget2.Object.Fields["Microsoft.VSTS.Common.StackRank"].Value, "2");

            //Publishing the second workitem again should no increase the stack rank
            _wordAdapter.Object.WorkItems.Remove(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Remove(workItemTarget1.Object);

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object, workItemSource2.Object }, false, configuration.Object);
            Assert.AreEqual(workItemTarget2.Object.Fields["Microsoft.VSTS.Common.StackRank"].Value, "2");
        }

        /// <summary>
        /// Make sure a link configured with direction <see cref="Direction.OtherToTfs" /> is published to TFS when publishing
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_Links_OtherToTfs_ShouldPublishLinks()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.OtherToTfs, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var workItemSource1 = CreateWorkItem(configurationItem, 1, "SourceValue1_");
            var workItemTarget1 = CreateWorkItem(configurationItem, 1, "TargetValue1_");

            _wordAdapter.Object.WorkItems.Add(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget1.Object);

            workItemSource1.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object }, false, configuration);

            // verify
            workItemSource1.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), It.IsAny<int[]>(), It.IsAny<string>(), It.IsAny<bool>()));
        }

        /// <summary>
        /// Make sure a link configured with direction <see cref="Direction.TfsToOther" /> is not published to TFS.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_Links_TfsToOther_ShouldNotPublishLinks()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.TfsToOther, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var workItemSource1 = CreateWorkItem(configurationItem, 1, "SourceValue1_");
            var workItemTarget1 = CreateWorkItem(configurationItem, 1, "TargetValue1_");

            _wordAdapter.Object.WorkItems.Add(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget1.Object);

            workItemSource1.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object }, false, configuration);

            // verify method is not called!
            workItemTarget1.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<bool>()), Times.Never());
        }

        /// <summary>
        /// Make sure a link configured with direction "PublishOnly" is published to TFS.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_Links_PublishOnly_ShouldPublishLinks()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.PublishOnly, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var workItemSource1 = CreateWorkItemLinkedItems(configurationItem, 1, "SourceValue1_");
            var workItemTarget1 = CreateWorkItem(configurationItem, 1, "TargetValue1_");

            _wordAdapter.Object.WorkItems.Add(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget1.Object);

            var workItemCollectionWithTargetWorkItem = new WorkItemCollection();
            workItemCollectionWithTargetWorkItem.Add(workItemTarget1.Object);
            workItemSource1.SetupGet(x => x.LinkedWorkItems).Returns(new Dictionary<string, IWorkItemCollection> { { linkConfiguration.LinkValueType, workItemCollectionWithTargetWorkItem } });

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object }, false, configuration);

            // verify
            workItemTarget1.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<bool>()));
        }

        /// <summary>
        /// Make sure a link configured with direction <see cref="Direction.SetInNewTfsWorkItem" /> is published to TFS for new work items
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_Links_SetInNewTfsWorkItem_ShouldPublishLinksToNewWorkItem()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.SetInNewTfsWorkItem, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var workItemSource1 = CreateWorkItemLinkedItems(configurationItem, 1, "SourceValue1_");
            var workItemTarget1 = CreateWorkItem(configurationItem, 1, "TargetValue1_");

            workItemTarget1.SetupGet(x => x.IsNew).Returns(true);

            _wordAdapter.Object.WorkItems.Add(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget1.Object);

            var workItemCollectionWithTargetWorkItem = new WorkItemCollection();
            workItemCollectionWithTargetWorkItem.Add(workItemTarget1.Object);
            workItemSource1.SetupGet(x => x.LinkedWorkItems).Returns(new Dictionary<string, IWorkItemCollection> { { linkConfiguration.LinkValueType, workItemCollectionWithTargetWorkItem } });

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object }, false, configuration);

            // verify
            workItemTarget1.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<bool>()));
        }

        /// <summary>
        /// Make sure a link configured with direction <see cref="Direction.SetInNewTfsWorkItem" /> is not published to TFS for old work items
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_Links_SetInNewTfsWorkItem_ShouldNotPublishLinksToOldWorkItem()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.SetInNewTfsWorkItem, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var workItemSource1 = CreateWorkItem(configurationItem, 1, "SourceValue1_");
            var workItemTarget1 = CreateWorkItem(configurationItem, 1, "TargetValue1_");

            workItemTarget1.SetupGet(x => x.IsNew).Returns(false);

            _wordAdapter.Object.WorkItems.Add(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget1.Object);

            workItemSource1.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object }, false, configuration);

            // verify
            workItemTarget1.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<bool>()), Times.Never());
        }

        /// <summary>
        /// Make sure links are not published for "GetOnly"
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_Links_GetOnly_ShouldNotPublishLinks()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.GetOnly, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var workItemSource1 = CreateWorkItem(configurationItem, 1, "SourceValue1_");
            var workItemTarget1 = CreateWorkItem(configurationItem, 1, "TargetValue1_");

            _wordAdapter.Object.WorkItems.Add(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget1.Object);

            workItemSource1.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object }, false, configuration);

            // verify
            workItemTarget1.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<bool>()), Times.Never());
        }

        /// <summary>
        /// Make sure links are refreshed for "GetOnly" for new work items.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_Links_GetOnly_ShouldRefreshForNewWorkItem()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.GetOnly, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            wordWorkItem.SetupGet(x => x.IsNew).Returns(true);

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            tfsWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            wordWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Exactly(1));
        }

        /// <summary>
        /// Test if a value of a bookmark is set for a staticvaluefield
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        [Ignore] // ALD: Configuration Variables no longer work.
        public void WorkItemSyncService_Refresh_BookmarksShouldBeSetForStaticValueField()
        {
            var configuration = CommonConfiguration.GetSimpleFieldConfigurationWithVariable("Requirement", Direction.PublishOnly, FieldValueType.DropDownList, "System.State");
            var configurationItem = configuration.GetConfigurationItems().Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            wordWorkItem.SetupGet(x => x.IsNew).Returns(true);

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            Assert.AreEqual(wordWorkItem.Object.Fields["VariableTest"].Value, CommonConfiguration.TestVariableText);
        }

        /// <summary>
        /// Make sure links are not refreshed for "GetOnly" on old work items.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_Links_GetOnly_ShouldNotRefreshForOldWorkItem()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.GetOnly, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            wordWorkItem.SetupGet(x => x.IsNew).Returns(false);

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            tfsWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });
            wordWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 999 } } });

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            wordWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<bool>()), Times.Never());
        }

        /// <summary>
        /// Make sure links for "GetOnly" are expanded if word and TFS are in sync.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_Links_GetOnly_ShouldExpandLinksIfWordAndTfsAreInSync()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.GetOnly, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            wordWorkItem.SetupGet(x => x.IsNew).Returns(false);

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            tfsWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });
            wordWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            wordWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Exactly(1));
        }

        /// <summary>
        /// Make sure links are refreshed for <see cref="Direction.OtherToTfs" />
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_Links_OtherToTfs_ShouldRefreshLinks()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.OtherToTfs, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            tfsWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            wordWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Exactly(1));
        }

        /// <summary>
        /// Make sure links are are refreshed for <see cref="Direction.TfsToOther" />
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_Links_TfsToOther_ShouldRefreshLinks()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.TfsToOther, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            tfsWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            wordWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Exactly(1));
        }

        /// <summary>
        /// Make sure links are not refreshed for "PublishOnly"
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_Links_PublishOnly_ShouldNotRefreshLinks()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.PublishOnly, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            tfsWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            wordWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<bool>()), Times.Never());
        }

        /// <summary>
        /// Make sure links are expanded for "PublishOnly" if the links are the same as on the server.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_Links_PublishOnly_ShouldNotRefreshLinksButExpandLinks()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.PublishOnly, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            tfsWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });
            wordWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            wordWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Exactly(1));
        }

        /// <summary>
        /// Make sure links are refreshed for "PublishOnly" on import.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_Links_PublishOnly_ShouldRefreshOnNewWorkItem()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.PublishOnly, "Parent", string.Empty, string.Empty, string.Empty, false);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            wordWorkItem.SetupGet(x => x.IsNew).Returns(true);

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            tfsWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 1 } } });

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            wordWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Exactly(1));
        }

        /// <summary>
        /// Make sure links are removed in word for link types that no longer exist on the TFS when refreshing.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Refresh_Links_ShouldRemoveLinkTypesNoLongerPresentOnTfs()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.PublishOnly, "Child", string.Empty, string.Empty, string.Empty, true);
            var configurationItem = configuration.GetConfigurationItems().Single();
            var linkConfiguration = configurationItem.Links.Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            wordWorkItem.SetupGet(x => x.Links).Returns(new Dictionary<IConfigurationLinkItem, int[]> { { linkConfiguration, new[] { 999 } } });

            // test
            _workItemSyncService.Refresh((IWorkItemSyncAdapter)_tfsServiceAdapter.Object, _wordAdapter.Object, new[] { wordWorkItem.Object }, configuration);

            // verify
            wordWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new int[] { }, "Child", It.IsAny<bool>()), Times.Exactly(1));
        }

        /// <summary>
        /// Make sure publishing an empty fields works
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_Links_PublishEmptyField_ShouldNotDoAnything()
        {
            var configuration = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.PublishOnly, "Child", string.Empty, string.Empty, string.Empty, true);
            var configurationItem = configuration.GetConfigurationItems().Single();

            var wordWorkItem = CreateWorkItem(configurationItem, 1, "WordWorkItem_");
            var tfsWorkItem = CreateWorkItem(configurationItem, 1, "TfsWorkItem_");

            _wordAdapter.Object.WorkItems.Add(wordWorkItem.Object);
            _tfsAdapter.Object.WorkItems.Add(tfsWorkItem.Object);

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { wordWorkItem.Object }, false, configuration);

            // verify
            tfsWorkItem.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), It.IsAny<int[]>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Never());
        }

        /// <summary>
        /// Even when setting a field to publish only, if it is a dropdown list at least the list of possible values must be updated.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_PublishOnly_DropDownList_ShouldRefreshDropDownList()
        {
            var configuration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.PublishOnly, FieldValueType.DropDownList, "System.State");

            var workItemSource1 = CreateWorkItem(configuration.GetConfigurationItems().First(), 1, "SourceValue1_");
            var workItemTarget1 = CreateWorkItem(configuration.GetConfigurationItems().First(), 1, "TargetValue1_");

            _wordAdapter.Object.WorkItems.Add(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget1.Object);

            workItemTarget1.Object.Fields["System.State"].AllowedValues = new List<string> { "A", "B" };
            workItemSource1.Object.Fields["System.State"].AllowedValues = new List<string> { "A" };
            workItemTarget1.Object.Fields["System.State"].Value = "B";
            workItemSource1.Object.Fields["System.State"].Value = "A";

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object }, false, configuration);
            FailOnUserInformation();

            // verify
            // List should be updated to AB, but value should remain A whereas tfs version has B
            Assert.IsTrue(workItemSource1.Object.Fields["System.State"].AllowedValues.Contains("B"));
            var field = workItemSource1.Object.Fields["System.State"];
            Mock.Get(field).VerifySet(x => x.Value = "A", Times.Exactly(2));
        }

        /// <summary>
        /// Even when setting a field to publish only, if it is a dropdown list at least the list of possible values must be updated.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_PublishOnly_DropDownList_ShouldNotRefreshDropDownListOfPublishOnlyFieldIfNewItemNotInList()
        {
            var configuration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.PublishOnly, FieldValueType.DropDownList, "System.State");

            var workItemSource1 = CreateWorkItem(configuration.GetConfigurationItems().First(), 1, "SourceValue1_");
            var workItemTarget1 = CreateWorkItem(configuration.GetConfigurationItems().First(), 1, "TargetValue1_");

            _wordAdapter.Object.WorkItems.Add(workItemSource1.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget1.Object);

            workItemTarget1.Object.Fields["System.State"].AllowedValues = new List<string> { "B", "C" };
            workItemSource1.Object.Fields["System.State"].AllowedValues = new List<string> { "A" };
            workItemTarget1.Object.Fields["System.State"].Value = "B";
            workItemSource1.Object.Fields["System.State"].Value = "A";

            // test
            _workItemSyncService.Publish(_wordAdapter.Object, (IWorkItemSyncAdapter)_tfsServiceAdapter.Object, new[] { workItemSource1.Object }, false, configuration);
            FailOnUserInformation();

            // verify
            Assert.IsFalse(workItemSource1.Object.Fields["System.State"].AllowedValues.Contains("B"));
            var field = workItemSource1.Object.Fields["System.State"];
            Mock.Get(field).VerifySet(x => x.Value = "A", Times.Exactly(1));
        }

        /// <summary>
        /// Check if the automatic link from requirement to issue is created when publishing.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_AutomaticLink_ShouldCreateLinkBetweenNewWorkItems()
        {
            var config = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.OtherToTfs, "Parent", "Issue", null, null, true);
            var tmpConfig = CommonConfiguration.GetSimpleFieldConfiguration("Issue", Direction.OtherToTfs, FieldValueType.HTML, "System.Description");

            var issue = CreateWorkItem(tmpConfig.GetConfigurationItems().First(), 0, "Issue");
            var tfsIssue = CreateWorkItem(tmpConfig.GetConfigurationItems().First(), 1, "TFSIssue");
            var requirement = CreateWorkItem(config.GetConfigurationItems().First(), 0, "Requirement");
            var tfsRequirement = CreateWorkItem(config.GetConfigurationItems().First(), 2, "TFSRequirement");

            _wordAdapter.Object.WorkItems.Add(issue.Object);
            _wordAdapter.Object.WorkItems.Add(requirement.Object);
            SetCreateNewQueue(_tfsAdapter, new[] { tfsIssue, tfsRequirement });
            FailOnUserInformation();

            // publish
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, _wordAdapter.Object.WorkItems, false, config.Merge(tmpConfig));

            tfsRequirement.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { 1 }, "Parent", true));
        }

        /// <summary>
        /// Check if the automatic link from requirement to the issue with the correct subtype (defined by My.Subtype) is created.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_AutomaticLink_ShouldCreateLinkBetweenNewWorkItemsWithCorrectSubtype()
        {
            var config = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.OtherToTfs, "Parent", "Issue", "My.Subtype", "CorrectSubtype", true);
            var tmpConfig = CommonConfiguration.GetSimpleFieldConfiguration("Issue", Direction.OtherToTfs, FieldValueType.PlainText, "My.Subtype");

            var correctIssue = CreateWorkItem(tmpConfig.GetConfigurationItems().First(), 0, "Issue");
            var correctTfsIssue = CreateWorkItem(tmpConfig.GetConfigurationItems().First(), 3, "TFSIssue");
            var wrongIssue = CreateWorkItem(tmpConfig.GetConfigurationItems().First(), 0, "Issue");
            var wrongTfsIssue = CreateWorkItem(tmpConfig.GetConfigurationItems().First(), 1, "TFSIssue");
            var requirement = CreateWorkItem(config.GetConfigurationItems().First(), 0, "Requirement");
            var tfsRequirement = CreateWorkItem(config.GetConfigurationItems().First(), 2, "TFSRequirement");
            correctIssue.Object.Fields["My.Subtype"].Value = "CorrectSubtype";
            wrongIssue.Object.Fields["My.Subtype"].Value = "WrongSubtype";

            // make sure the wrong issue is direcly before the requirement
            _wordAdapter.Object.WorkItems.Add(correctIssue.Object);
            _wordAdapter.Object.WorkItems.Add(wrongIssue.Object);
            _wordAdapter.Object.WorkItems.Add(requirement.Object);

            SetCreateNewQueue(_tfsAdapter, new[] { correctTfsIssue, wrongTfsIssue, tfsRequirement });

            // publish
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, _wordAdapter.Object.WorkItems, false, config.Merge(tmpConfig));
            FailOnUserInformation();

            tfsRequirement.Verify(x => x.AddLinks(It.IsAny<IWorkItemSyncAdapter>(), new[] { correctTfsIssue.Object.Id }, "Parent", true));
        }

        /// <summary>
        /// Check if the automatic link from requirement to the issue with the correct subtype (defined by My.Subtype) is created.
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_AutomaticLink_ShouldNagIfSubtypeFieldIsDefinedButDoesNotExist()
        {
            var config = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.OtherToTfs, "Parent", "Issue", "My.Subtype", "CorrectSubtype", true);
            var tmpConfig = CommonConfiguration.GetSimpleFieldConfiguration("Issue", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");

            var correctIssue = CreateWorkItem(tmpConfig.GetConfigurationItems().First(), 0, "Issue");
            var correctTfsIssue = CreateWorkItem(tmpConfig.GetConfigurationItems().First(), 3, "TFSIssue");
            var requirement = CreateWorkItem(config.GetConfigurationItems().First(), 0, "Requirement");
            var tfsRequirement = CreateWorkItem(config.GetConfigurationItems().First(), 2, "TFSRequirement");

            // make sure the wrong issue is direcly before the requirement
            _wordAdapter.Object.WorkItems.Add(correctIssue.Object);
            _wordAdapter.Object.WorkItems.Add(requirement.Object);

            SetCreateNewQueue(_tfsAdapter, new[] { correctTfsIssue, tfsRequirement });

            // publish
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, _wordAdapter.Object.WorkItems, false, config.Merge(tmpConfig));
            ExpectUserInformation(3);
        }

        /// <summary>
        /// Calling "Publish Selected" on a work item where the auto link target is new and not also published should add user information
        /// </summary>
        [TestMethod]
        public void WorkItemSyncService_Publish_AutomaticLink_ShouldNagIfPublishSelectedAndTargetIsNewAndNotPublished()
        {
            var config = CommonConfiguration.GetSimpleLinkConfiguration("Requirement", Direction.OtherToTfs, "Parent", "Issue", "My.Subtype", "CorrectSubtype", true);
            var tmpConfig = CommonConfiguration.GetSimpleFieldConfiguration("Issue", Direction.OtherToTfs, FieldValueType.PlainText, "My.Subtype");

            var correctIssue = CreateWorkItem(tmpConfig.GetConfigurationItems().First(), 0, "Issue");
            var requirement = CreateWorkItem(config.GetConfigurationItems().First(), 0, "Requirement");
            var tfsRequirement = CreateWorkItem(config.GetConfigurationItems().First(), 2, "TFSRequirement");
            correctIssue.Object.Fields["My.Subtype"].Value = "CorrectSubtype";

            // make sure the wrong issue is direcly before the requirement
            _wordAdapter.Object.WorkItems.Add(correctIssue.Object);
            _wordAdapter.Object.WorkItems.Add(requirement.Object);

            SetCreateNewQueue(_tfsAdapter, new[] { tfsRequirement });

            // publish selected (requirement)
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, new[] { requirement.Object }, false, config.Merge(tmpConfig));
            ExpectUserInformation(2);
        }

        ///  <summary>
        /// Test if builds for a test case can be retrieved
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        [Ignore] // Needs linked build
        public void WorkItemSyncService_GetBuildInformationForTestCaseDetail_ShouldReturnBuildsForTestCaseWithBuildInformation()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = adapter.TestManagement.TestCases.Find(CommonConfiguration.TestCaseIdWithBuildInformation);
            var testSuite = adapter.TestManagement.TestSuites.Find(CommonConfiguration.TestSuiteIdWithBuildInformation);
            var testPlan = adapter.TestManagement.TestPlans.Find(CommonConfiguration.TestPlanIdWithBuildInformation);
            var tfsTestPlan = new TfsTestPlan(testPlan);

            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase, new TfsTestSuite((IStaticTestSuite)testSuite, tfsTestPlan)));
            var builds = adapter.GetAllBuildsForTestCase(tfsTestCase, true);

            Assert.IsNotNull(builds);
            Assert.IsTrue(builds.Any());
        }

        ///  <summary>
        /// Test if builds for a test case can be retrieved
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_GetBuildInformationForTestCaseDetail_ShouldNotReturnBuildsForTestCaseWithOutBuildInformation()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = adapter.TestManagement.TestCases.Find(CommonConfiguration.TestCaseId);
            var testSuite = adapter.TestManagement.TestSuites.Find(CommonConfiguration.TestSuiteId);
            var testPlan = adapter.TestManagement.TestPlans.Find(CommonConfiguration.TestPlanId);
            var tfsTestPlan = new TfsTestPlan(testPlan);

            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase, new TfsTestSuite((IStaticTestSuite)testSuite, tfsTestPlan)));
            var builds = adapter.GetAllBuildsForTestCase(tfsTestCase, true);

            Assert.IsNotNull(builds);
            Assert.IsFalse(builds.Any());
        }

        ///  <summary>
        /// Test if linked bugs are returned when no filters are set
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_GetLinkedWorkItemsForTestCase_ShouldReturnLinkedBug_ForValidTestCasesAndNoFilters()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = adapter.TestManagement.TestCases.Find(CommonConfiguration.TestCaseId);

            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase));
            var workItemTypes = new List<string>();
            workItemTypes.Add("Bug");

            var testCases = new List<ITfsTestCaseDetail>();
            testCases.Add(tfsTestCase);
            var linkFilter = new WorkItemLinkFilter();
            var filterOption = new FilterOption();
            var linkedWorkItems = adapter.GetAllLinkedWorkItemsForTestCases(testCases, workItemTypes, linkFilter, filterOption);

            //List should contain work items, all of type "B_ug"
            Assert.IsNotNull(linkedWorkItems);
            Assert.IsTrue(linkedWorkItems.Any());
            Assert.IsTrue(linkedWorkItems.TrueForAll(x => x.Type.Name.Equals("Bug")));
        }

        ///  <summary>
        /// Test the distinct filter
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_GetLinkedWorkItemsForTestCase_ShouldReturnLinkedBug_ForValidTestCasesAndDistinctFilter()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            // ReSharper disable once CSharpWarnings::CS0618
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = adapter.TestManagement.TestCases.Find(CommonConfiguration.TestCaseId);

            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase));
            var workItemTypes = new List<string>();
            workItemTypes.Add("Bug");

            var testCases = new List<ITfsTestCaseDetail>();
            testCases.Add(tfsTestCase);
            var linkFilter = new WorkItemLinkFilter();

            var filterOption = new FilterOption();
            filterOption.Distinct = true;
            var linkedWorkItems = adapter.GetAllLinkedWorkItemsForTestCases(testCases, workItemTypes, linkFilter, filterOption);

            //List should contain work items, all of type "B_ug"
            Assert.IsNotNull(linkedWorkItems);
            Assert.IsTrue(linkedWorkItems.Any());
            Assert.IsTrue(linkedWorkItems.TrueForAll(x => x.Type.Name.Equals("Bug")));
        }

        ///  <summary>
        /// An empty link filter should return the bugs anyway
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_GetLinkedWorkItemsForTestCase_ShouldReturnLinkedBug_ForValidTestCasesAndDistinctFilterAndEmptyIncludeFilter()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            // ReSharper disable once CSharpWarnings::CS0618
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = adapter.TestManagement.TestCases.Find(CommonConfiguration.TestCaseId);

            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase));
            var workItemTypes = new List<string>();
            workItemTypes.Add("Bug");

            var testCases = new List<ITfsTestCaseDetail>();
            testCases.Add(tfsTestCase);
            var linkFilter = new WorkItemLinkFilter();
            linkFilter.FilterType = FilterType.Include;

            var filterOption = new FilterOption();
            filterOption.Distinct = true;
            var linkedWorkItems = adapter.GetAllLinkedWorkItemsForTestCases(testCases, workItemTypes, linkFilter, filterOption);

            //List should contain work items, all of type "B_ug"
            Assert.IsNotNull(linkedWorkItems);
            Assert.IsTrue(linkedWorkItems.Any());
            Assert.IsTrue(linkedWorkItems.TrueForAll(x => x.Type.Name.Equals("Bug")));
        }

        /// <summary>
        /// Test if the include filter works
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_GetLinkedWorkItemsForTestCase_ShouldReturnLinkedBug_ForValidTestCasesAndDistinctFilterAndCorrectIncludeFilter()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = adapter.TestManagement.TestCases.Find(CommonConfiguration.TestCaseId);

            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase));
            var workItemTypes = new List<string>();
            workItemTypes.Add("Bug");

            var testCases = new List<ITfsTestCaseDetail>();
            testCases.Add(tfsTestCase);
            var linkFilter = new WorkItemLinkFilter();
            linkFilter.FilterType = FilterType.Include;

            var includeFilter = new Filter();
            includeFilter.LinkType = "Microsoft.VSTS.Common.TestedBy-Forward";
            linkFilter.Filters.Add(includeFilter);

            var filterOption = new FilterOption();
            filterOption.Distinct = true;
            var linkedWorkItems = adapter.GetAllLinkedWorkItemsForTestCases(testCases, workItemTypes, linkFilter, filterOption);

            //List should contain work items, all of type "B_ug"
            Assert.IsNotNull(linkedWorkItems);
            Assert.IsTrue(linkedWorkItems.Any());
            Assert.IsTrue(linkedWorkItems.TrueForAll(x => x.Type.Name.Equals("Bug")));
        }

        /// <summary>
        /// Test the exclude filter. Filter all #bugs
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_GetLinkedWorkItemsForTestCase_ShouldNotReturnLinkedBug_ForValidTestCasesAndDistinctFilterAndCorrectExcludeFilter()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = adapter.TestManagement.TestCases.Find(CommonConfiguration.TestCaseId);

            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase));
            var workItemTypes = new List<string> {"Bug"};
            var testCases = new List<ITfsTestCaseDetail>();
            testCases.Add(tfsTestCase);
            var linkFilter = new WorkItemLinkFilter();
            linkFilter.FilterType = FilterType.Exclude;

            var includeFilter = new Filter();
            includeFilter.LinkType = "Microsoft.VSTS.Common.TestedBy-Forward";
            linkFilter.Filters.Add(includeFilter);

            var filterOption = new FilterOption();
            filterOption.Distinct = true;
            var linkedWorkItems = adapter.GetAllLinkedWorkItemsForTestCases(testCases, workItemTypes, linkFilter, filterOption);

            //List should not contain any work items
            Assert.IsNotNull(linkedWorkItems);
            Assert.IsFalse(linkedWorkItems.Any());
        }

        /// <summary>
        /// Test the exclude filter. Filter all #bugs
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_GetLinkedWorkItemsForTestCase_ShouldReturnEmptyList_ForValidTestCasesAndWrongFilter()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = adapter.TestManagement.TestCases.Find(CommonConfiguration.TestCaseId);

            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase));
            var workItemTypes = new List<string>();
            workItemTypes.Add("AWorkItemThatDoesNotExist");

            var testCases = new List<ITfsTestCaseDetail>();
            testCases.Add(tfsTestCase);
            var linkFilter = new WorkItemLinkFilter();
            linkFilter.FilterType = FilterType.Include;

            var includeFilter = new Filter();
            includeFilter.LinkType = "Microsoft.VSTS.Common.TestedBy-Forward";
            linkFilter.Filters.Add(includeFilter);

            var filterOption = new FilterOption();
            filterOption.Distinct = true;
            var linkedWorkItems = adapter.GetAllLinkedWorkItemsForTestCases(testCases, workItemTypes, linkFilter, filterOption);

            //List should not contain any work items
            Assert.IsNotNull(linkedWorkItems);
            Assert.IsFalse(linkedWorkItems.Any());
        }

        /// <summary>
        /// Test the exclude filter. Filter all #bugs
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_GetLinkedWorkItemsForTestCase_ShouldReturnEmptyList_ForValidTestCasesAndWrongLinkTypeFilter()
        {
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.TestCaseId }));
            var testCase = adapter.TestManagement.TestCases.Find(CommonConfiguration.TestCaseId);

            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase));
            var workItemTypes = new List<string>();
            workItemTypes.Add("Bug");

            var testCases = new List<ITfsTestCaseDetail>();
            testCases.Add(tfsTestCase);
            var linkFilter = new WorkItemLinkFilter();
            linkFilter.FilterType = FilterType.Include;

            var includeFilter = new Filter();
            includeFilter.LinkType = "ALinkTypeThatDoesNotExists";
            linkFilter.Filters.Add(includeFilter);

            var filterOption = new FilterOption();
            filterOption.Distinct = true;
            var linkedWorkItems = adapter.GetAllLinkedWorkItemsForTestCases(testCases, workItemTypes, linkFilter, filterOption);

            //List should not contain any work items
            Assert.IsNotNull(linkedWorkItems);
            Assert.IsFalse(linkedWorkItems.Any());
        }

        ///  <summary>
        /// Test if linked bugs are returned when no filters are set
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_GetLinkedWorkItemsForTestResult_ShouldReturnLinkedBug_ForValidTestCasesAndNoFilters()
        {
            var testCaseId = CommonConfiguration.TestCaseIdWithTestResultLinkedToBug;
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            // ReSharper disable once CSharpWarnings::CS0618
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { testCaseId }));
            var testResults = adapter.TestManagement.TestResults.ByTestId(testCaseId);
            var latestFailedTestResult = testResults.OrderByDescending(x => x.DateCompleted).First(x => x.Outcome == TestOutcome.Failed);
            var latestSuccededTestResult = testResults.OrderByDescending(x => x.DateCompleted).First(x => x.Outcome == TestOutcome.Passed);

            var failedTfsTestResult = new TfsTestResultDetail(new TfsTestResult(latestFailedTestResult, testCaseId), latestFailedTestResult.TestRunId);
            var succededTfsTestResult = new TfsTestResultDetail(new TfsTestResult(latestSuccededTestResult, testCaseId), latestSuccededTestResult.TestRunId);

            var workItemTypes = new List<string>();
            workItemTypes.Add("Bug");

            var linkFilter = new WorkItemLinkFilter();
            var filterOption = new FilterOption();

            var linkedWorkItemsForFailedTest = adapter.GetAllLinkedWorkItemsForTestResult(failedTfsTestResult, workItemTypes, linkFilter, filterOption);
            var linkedWorkItemsForSuccededTest = adapter.GetAllLinkedWorkItemsForTestResult(succededTfsTestResult, workItemTypes, linkFilter, filterOption);

            //List should contain work items, all of type "B_ug"
            Assert.IsNotNull(linkedWorkItemsForFailedTest);
            Assert.IsTrue(linkedWorkItemsForFailedTest.Any());
            Assert.IsTrue(linkedWorkItemsForFailedTest.TrueForAll(x => x.Type.Name.Equals("Bug")));

            //List should not contain work items
            Assert.IsNotNull(linkedWorkItemsForSuccededTest);
            Assert.IsFalse(linkedWorkItemsForSuccededTest.Any());
        }

        ///  <summary>
        /// Test if linked bugs are returned when no filters are set
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void WorkItemSyncService_CompareLinkedWorkItemsForTestResultAndTestCases_ShouldBeEqual_ForValidTestCasesAndNoFilters()
        {
            var testCaseId = CommonConfiguration.TestCaseIdWithTestResultLinkedToBug;
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));

            var adapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, serverConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { testCaseId }));

            //Get the work items for the test result
            var testResults = adapter.TestManagement.TestResults.ByTestId(testCaseId);
            var latestFailedTestResult = testResults.OrderByDescending(x => x.DateCompleted).First(x => x.Outcome == TestOutcome.Failed);
            var latestSuccededTestResult = testResults.OrderByDescending(x => x.DateCompleted).First(x => x.Outcome == TestOutcome.Passed);

            var failedTfsTestResult = new TfsTestResultDetail(new TfsTestResult(latestFailedTestResult, testCaseId), latestFailedTestResult.TestRunId);
            var succededTfsTestResult = new TfsTestResultDetail(new TfsTestResult(latestSuccededTestResult, CommonConfiguration.TestCaseId), latestSuccededTestResult.TestRunId);

            var workItemTypes = new List<string>();
            workItemTypes.Add("Bug");

            var linkFilter = new WorkItemLinkFilter();
            var filterOption = new FilterOption();

            var linkedWorkItemsForFailedTest = adapter.GetAllLinkedWorkItemsForTestResult(failedTfsTestResult, workItemTypes, linkFilter, filterOption);
            var linkedWorkItemsForSuccededTest = adapter.GetAllLinkedWorkItemsForTestResult(succededTfsTestResult, workItemTypes, linkFilter, filterOption);

            //Get the linked workitem for the corresponding test case
            var testCase = adapter.TestManagement.TestCases.Find(testCaseId);

            var testCaseList = new List<ITfsTestCaseDetail>();
            var tfsTestCase = new TfsTestCaseDetail(new TfsTestCase(testCase));

            testCaseList.Add(tfsTestCase);

            var linkedWorkItemsForTestCase = adapter.GetAllLinkedWorkItemsForTestCases(testCaseList, workItemTypes, linkFilter, filterOption);

            Assert.IsTrue(linkedWorkItemsForTestCase.Any(l1 => linkedWorkItemsForFailedTest.Any(l2 => l1.Id == l2.Id)));

            //List should contain work items, all of type "Bug"
            Assert.IsNotNull(linkedWorkItemsForFailedTest);
            Assert.IsTrue(linkedWorkItemsForFailedTest.Any());
            Assert.IsTrue(linkedWorkItemsForFailedTest.TrueForAll(x => x.Type.Name.Equals("Bug")));

            //List should not contain work items
            Assert.IsNotNull(linkedWorkItemsForSuccededTest);
            Assert.IsFalse(linkedWorkItemsForSuccededTest.Any());
        }

        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldSetOLEMarkerFromTwoSourcesToOneDestinationWithCustomValues()
        {
            // Arrange
            var oleMarkerField = "System.Title";
            var systemDescription = "System.Description";
            var cMMImpactAssessmentHtml = "Microsoft.VSTS.CMMI.ImpactAssessmentHtml";
            var customDesc = "customDesc";
            var customAnalysis = "customAnalysis";
            var configuration = CommonConfiguration.Configuration.GetConfigurationItems().First(x => x.WorkItemType == "Requirement").Clone();
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(systemDescription, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, customDesc, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(cMMImpactAssessmentHtml, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, customAnalysis, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            var workItemSource = CreateWorkItem(configuration, 1, "SourceValue_");
            var workItemTarget = CreateWorkItem(configuration, 1, "TargetValue_");

            var descripionField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            descripionField.SetupAllProperties();
            descripionField.SetupGet(x => x.ReferenceName).Returns(systemDescription);
            descripionField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, descripionField);

            var impactAssessmentHtmlField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            impactAssessmentHtmlField.SetupAllProperties();
            impactAssessmentHtmlField.SetupGet(x => x.ReferenceName).Returns(cMMImpactAssessmentHtml);
            impactAssessmentHtmlField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, impactAssessmentHtmlField);

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // Act
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, new[] { workItemSource.Object }, false, CommonConfiguration.Configuration);

            // Assert
            // ReSharper disable once UseStringInterpolation (MIS:Relase Build does not support C#6 features)
            var shouldValue = string.Format("{0}, {1}",customDesc, customAnalysis);
            Assert.AreEqual(workItemTarget.Object.Fields[oleMarkerField].Value, shouldValue);
        }

        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldSetOLEMarkerFromTwoSourcesToOneDestinationWithDefaultValues()
        {
            // Arrange
            var description = "Description";
            var analysis = "Analysis";
            var oleMarkerField = "System.Title";
            var systemDescription = "System.Description";
            var cMMImpactAssessmentHtml = "Microsoft.VSTS.CMMI.ImpactAssessmentHtml";
            var configuration = CommonConfiguration.Configuration.GetConfigurationItems().First(x => x.WorkItemType == "Requirement").Clone();
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(systemDescription, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, null, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(cMMImpactAssessmentHtml, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, null, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            var workItemSource = CreateWorkItem(configuration, 1, "SourceValue_");

            var descripionField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            descripionField.SetupAllProperties();
            descripionField.SetupGet(x => x.ReferenceName).Returns(systemDescription);
            descripionField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, descripionField);

            var impactAssessmentHtmlField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            impactAssessmentHtmlField.SetupAllProperties();
            impactAssessmentHtmlField.SetupGet(x => x.ReferenceName).Returns(cMMImpactAssessmentHtml);
            impactAssessmentHtmlField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, impactAssessmentHtmlField);

            var workItemTarget = CreateWorkItem(configuration, 1, "TargetValue_");

            var targetDescripionField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            targetDescripionField.SetupAllProperties();
            targetDescripionField.SetupGet(x => x.Name).Returns(description);
            targetDescripionField.SetupGet(x => x.ReferenceName).Returns(systemDescription);
            AddOrUpdateField(configuration, workItemTarget.Object.Fields, targetDescripionField);

            var targetimpactAssessmentHtmlField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            targetimpactAssessmentHtmlField.SetupAllProperties();
            targetimpactAssessmentHtmlField.SetupGet(x => x.Name).Returns(analysis);
            targetimpactAssessmentHtmlField.SetupGet(x => x.ReferenceName).Returns(cMMImpactAssessmentHtml);
            AddOrUpdateField(configuration, workItemTarget.Object.Fields, targetimpactAssessmentHtmlField);

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // Act
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, new[] { workItemSource.Object }, false, CommonConfiguration.Configuration);

            // Assert
            // ReSharper disable once UseStringInterpolation (MIS:Relase Build does not support C#6 features)
            var shouldValue = string.Format("{0}, {1}", description, analysis);
            Assert.AreEqual(workItemTarget.Object.Fields[oleMarkerField].Value, shouldValue);
        }

        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldSetOLEMarkerFromTwoSourcesToTwoDestinationsWithDefaultValues()
        {
            // Arrange
            var descriptionOleMarkerField = "System.Title";
            var impactAssessmentHtmlOleMarkerField = "System.Reason";
            var systemDescription = "System.Description";
            var cMMImpactAssessmentHtml = "Microsoft.VSTS.CMMI.ImpactAssessmentHtml";
            var configuration = CommonConfiguration.Configuration.GetConfigurationItems().First(x => x.WorkItemType == "Requirement").Clone();
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem("System.Reason", string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, null, null, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(systemDescription, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, descriptionOleMarkerField, null, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(cMMImpactAssessmentHtml, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, impactAssessmentHtmlOleMarkerField, null, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            var workItemSource = CreateWorkItem(configuration, 1, "SourceValue_");

            var descripionField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            descripionField.SetupAllProperties();
            descripionField.SetupGet(x => x.ReferenceName).Returns(systemDescription);
            descripionField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, descripionField);

            var impactAssessmentHtmlField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            impactAssessmentHtmlField.SetupAllProperties();
            impactAssessmentHtmlField.SetupGet(x => x.ReferenceName).Returns(cMMImpactAssessmentHtml);
            impactAssessmentHtmlField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, impactAssessmentHtmlField);

            var workItemTarget = CreateWorkItem(configuration, 1, "TargetValue_");

            var targetDescripionField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            targetDescripionField.SetupAllProperties();
            targetDescripionField.SetupGet(x => x.Name).Returns("Description");
            targetDescripionField.SetupGet(x => x.ReferenceName).Returns(systemDescription);
            AddOrUpdateField(configuration, workItemTarget.Object.Fields, targetDescripionField);

            var targetimpactAssessmentHtmlField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            targetimpactAssessmentHtmlField.SetupAllProperties();
            targetimpactAssessmentHtmlField.SetupGet(x => x.Name).Returns("Analysis");
            targetimpactAssessmentHtmlField.SetupGet(x => x.ReferenceName).Returns(cMMImpactAssessmentHtml);
            AddOrUpdateField(configuration, workItemTarget.Object.Fields, targetimpactAssessmentHtmlField);

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // Act
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, new[] { workItemSource.Object }, false, CommonConfiguration.Configuration);

            // Assert
            Assert.AreEqual(workItemTarget.Object.Fields[descriptionOleMarkerField].Value, "Description");
            Assert.AreEqual(workItemTarget.Object.Fields[impactAssessmentHtmlOleMarkerField].Value, "Analysis");
        }

        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldSetOLEMarkerFromTwoSourcesToOneDestinationY()
        {
            // Arrange
            var oleMarkerField = "System.Title";
            var oleMarkerValue = "Y";
            var systemDescription = "System.Description";
            var cMMImpactAssessmentHtml = "Microsoft.VSTS.CMMI.ImpactAssessmentHtml";
            var configuration = CommonConfiguration.Configuration.GetConfigurationItems().First(x => x.WorkItemType == "Requirement").Clone();
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(systemDescription, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, oleMarkerValue, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(cMMImpactAssessmentHtml, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, oleMarkerValue, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            var workItemSource = CreateWorkItem(configuration, 1, "SourceValue_");
            var workItemTarget = CreateWorkItem(configuration, 1, "TargetValue_");

            var descripionField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            descripionField.SetupAllProperties();
            descripionField.SetupGet(x => x.ReferenceName).Returns(systemDescription);
            descripionField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, descripionField);

            var impactAssessmentHtmlField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            impactAssessmentHtmlField.SetupAllProperties();
            impactAssessmentHtmlField.SetupGet(x => x.ReferenceName).Returns(cMMImpactAssessmentHtml);
            impactAssessmentHtmlField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, impactAssessmentHtmlField);

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // Act
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, new[] { workItemSource.Object }, false, CommonConfiguration.Configuration);

            // Assert
            Assert.AreEqual(workItemTarget.Object.Fields[oleMarkerField].Value, oleMarkerValue);
        }

        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldSetOLEMarkerFromTwoSourcesWhenOnlyOneSourceContainsOleObject()
        {
            // Arrange
            var oleMarkerField = "System.Title";
            var oleMarkerValue = "Y";
            var systemDescription = "System.Description";
            var cMMImpactAssessmentHtml = "Microsoft.VSTS.CMMI.ImpactAssessmentHtml";
            var configuration = CommonConfiguration.Configuration.GetConfigurationItems().First(x => x.WorkItemType == "Requirement").Clone();
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(systemDescription, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, oleMarkerValue, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(cMMImpactAssessmentHtml, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, oleMarkerValue, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            var workItemSource = CreateWorkItem(configuration, 1, "SourceValue_");
            var workItemTarget = CreateWorkItem(configuration, 1, "TargetValue_");

            var descripionField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            descripionField.SetupAllProperties();
            descripionField.SetupGet(x => x.ReferenceName).Returns(systemDescription);
            descripionField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, descripionField);

            var impactAssessmentHtmlField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            impactAssessmentHtmlField.SetupAllProperties();
            impactAssessmentHtmlField.SetupGet(x => x.ReferenceName).Returns(cMMImpactAssessmentHtml);
            impactAssessmentHtmlField.SetupGet(x => x.ContainsOleObject).Returns(false);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, impactAssessmentHtmlField);

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // Act
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, new[] { workItemSource.Object }, false, CommonConfiguration.Configuration);

            // Assert
            Assert.AreEqual(workItemTarget.Object.Fields[oleMarkerField].Value, oleMarkerValue);
        }

        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldSetOLEMarkerFromOneSourceToOneDestinationY()
        {
            // Arrange
            var oleMarkerField = "System.Title";
            var oleMarkerValue = "Y";
            var cMMImpactAssessmentHtml = "Microsoft.VSTS.CMMI.ImpactAssessmentHtml";
            var configuration = CommonConfiguration.Configuration.GetConfigurationItems().First(x => x.WorkItemType == "Requirement").Clone();
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(cMMImpactAssessmentHtml, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, oleMarkerValue, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            var workItemSource = CreateWorkItem(configuration, 1, "SourceValue_");
            var workItemTarget = CreateWorkItem(configuration, 1, "TargetValue_");

            var impactAssessmentHtmlField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            impactAssessmentHtmlField.SetupAllProperties();
            impactAssessmentHtmlField.SetupGet(x => x.ReferenceName).Returns(cMMImpactAssessmentHtml);
            impactAssessmentHtmlField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, impactAssessmentHtmlField);

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // Act
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, new[] { workItemSource.Object }, false, CommonConfiguration.Configuration);

            // Assert
            Assert.AreEqual(workItemTarget.Object.Fields[oleMarkerField].Value, oleMarkerValue);
        }

        [TestMethod]
        public void WorkItemSyncService_Publish_ShouldSetOLEMarkerFromOneSourceToOneDestination()
        {
            // Arrange
            var oleMarkerField = "System.Title";
            var oleMarkerValue = "customDesc";
            var systemDescription = "System.Description";
            var configuration = CommonConfiguration.Configuration.GetConfigurationItems().First(x => x.WorkItemType == "Requirement").Clone();
            AddOrUpdateConfigurationFieldItem(configuration.FieldConfigurations, new ConfigurationFieldItem(systemDescription, string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 1, 1, string.Empty, false, HandleAsDocumentType.OleOnDemand, oleMarkerField, oleMarkerValue, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            var workItemSource = CreateWorkItem(configuration, 1, "SourceValue_");
            var workItemTarget = CreateWorkItem(configuration, 1, "TargetValue_");

            var impactAssessmentHtmlField = new Mock<IField> { DefaultValue = DefaultValue.Mock };
            impactAssessmentHtmlField.SetupAllProperties();
            impactAssessmentHtmlField.SetupGet(x => x.ReferenceName).Returns(systemDescription);
            impactAssessmentHtmlField.SetupGet(x => x.ContainsOleObject).Returns(true);
            AddOrUpdateField(configuration, workItemSource.Object.Fields, impactAssessmentHtmlField);

            _wordAdapter.Object.WorkItems.Add(workItemSource.Object);
            _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

            // Act
            _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, new[] { workItemSource.Object }, false, CommonConfiguration.Configuration);

            // Assert
            Assert.AreEqual(workItemTarget.Object.Fields[oleMarkerField].Value, oleMarkerValue);
        }

        // Test don't working, should be fix
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        [Ignore]
        public void WorkItemSyncService_Publish_ShouldPublishDescriptionFieldOnlyWithImageWithoutText()
        {
            // Arrange
            var configuration = CommonConfiguration.GetCMMIRequirement();
            var configurationItem = configuration.GetConfigurationItems().Single();

            //WORK A ROUND : Build path to docx file
            //Document path is folder path + assembly name + relative path to file
            //Problem is that assembly name is not same as folder neme (".Unit" must be removed from name).
            var folderSourcesPath = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.TestDir));
            var projectName = Assembly.GetExecutingAssembly().GetName().Name.Replace(".Unit", "");
            var relativePath = "Test Data\\DocWithTable.docx";
            if (folderSourcesPath != null)
            {
                var documentPath = Path.Combine(folderSourcesPath, projectName, relativePath);

                var wordApplication = new Application();
                var wordDocument = wordApplication.Application.Documents.Open(documentPath, true);
                var table = wordDocument.Tables[1];

                var workItemSource = new WordTableWorkItem(table, "Requirement", CommonConfiguration.Configuration, configurationItem);
                var workItemTarget = CreateTargetWorkItem(workItemSource);

                _wordAdapter.Object.WorkItems.Add(workItemSource);
                _tfsAdapter.Object.WorkItems.Add(workItemTarget.Object);

                // Act
                _workItemSyncService.Publish(_wordAdapter.Object, _tfsAdapter.Object, new[] { workItemSource }, false, CommonConfiguration.Configuration);

                // Assert
                Assert.IsFalse(workItemTarget.Object.Fields["System.Description"].Value == string.Empty);

                wordDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
                wordApplication.Quit(true);
            }
        }

        #endregion Tests

        #region Helper members

        private static Mock<IWorkItem> CreateTargetWorkItem(WordTableWorkItem workItemSource)
        {
            var workItem = new Mock<IWorkItem> { DefaultValue = DefaultValue.Mock };
            workItem.SetupAllProperties();
            var fieldCollection = new FieldCollection();

            foreach (var fieldConfiguration in workItemSource.Configuration.FieldConfigurations)
            {
                var field = new Mock<IField> { DefaultValue = DefaultValue.Mock };

                var referenceName = fieldConfiguration.ReferenceFieldName;
                field.SetupAllProperties();

                field.SetupGet(x => x.ReferenceName).Returns(referenceName);
                field.SetupGet(x => x.IsEditable).Returns(true);
                field.SetupGet(x => x.Configuration).Returns(fieldConfiguration);

                if (referenceName == "System.Description")
                {
                    field.Setup(f => f.CompareValue(It.IsAny<string>(), It.IsAny<bool>())).Returns(false);
                    field.Object.Value = string.Empty;
                }
                else
                {
                    field.Setup(f => f.CompareValue(It.IsAny<string>(), It.IsAny<bool>())).Returns(true);
                    field.Object.Value = workItemSource.Fields[referenceName].Value;
                }

                fieldCollection.Add(field.Object);
            }

            workItem.SetupGet(x => x.Id).Returns(workItemSource.Id);
            workItem.SetupGet(x => x.Fields).Returns(fieldCollection);
            workItem.SetupGet(x => x.WorkItemType).Returns(workItemSource.WorkItemType);
            workItem.SetupGet(x => x.Revision).Returns(1);
            workItem.Setup(x => x.GetFieldRevision(It.IsAny<string>())).Returns(1);
            workItem.Setup(x => x.GetWorkItemByRevision(It.IsAny<int>())).Returns(fieldCollection);
            workItem.SetupGet(x => x.Configuration).Returns(workItemSource.Configuration);
            workItem.SetupGet(x => x.IsNew).Returns(workItem.Object.Id == 0);
            workItem.Setup(x => x.IsSameWorkItem(It.IsAny<IWorkItem>())).Returns((IWorkItem wi) => (wi.Id != 0 && wi.Id == workItem.Object.Id) || (wi.Id == 0 && wi == workItem.Object));
            if (workItem.Object.Fields.Contains("System.Title"))
            {
                workItem.SetupGet(x => x.Title).Returns(workItem.Object.Fields["System.Title"].Value);
            }
            else
            {
                workItem.SetupGet(x => x.Title).Returns("Value not set");
            }

            return workItem;
        }

        private static void SetCreateNewQueue(Mock<IWorkItemSyncAdapter> adapter, IEnumerable<Mock<IWorkItem>> workItems)
        {
            var index = 0;

            adapter.Setup(x => x.CreateNewWorkItem(It.IsAny<IConfigurationItem>())).Returns(
                () =>
                {
                    var workItem = workItems.ElementAt(index);
                    index++;
                    adapter.Object.WorkItems.Add(workItem.Object);
                    return workItem.Object;
                });
        }

        /// <summary>
        /// Get the infostorage and fail if messages come form wordtotfs
        /// </summary>
        private static void FailOnUserInformation()
        {
            var informationStorage = SyncServiceFactory.GetService<IInfoStorageService>();
            var error = informationStorage.UserInformation.FirstOrDefault(x => x.Type == UserInformationType.Error);
            if (error != null)
            {
                Assert.Fail(error.Explanation);
            }
        }

        // ReSharper disable once UnusedParameter.Local
        private static void ExpectUserInformation(int expectedCount)
        {
            var informationStorage = SyncServiceFactory.GetService<IInfoStorageService>();
            var count = informationStorage.UserInformation.Count;
            if (count != expectedCount)
            {
                Assert.Fail("There are not exactly as many user information entries as expected");
            }
        }

        private static void AddOrUpdateConfigurationFieldItem(IList<IConfigurationFieldItem> configuratins, ConfigurationFieldItem configurationFieldItem)
        {
            var currentConfigurationFieldItem = configuratins.FirstOrDefault(f => f.ReferenceFieldName == configurationFieldItem.ReferenceFieldName);
            if (currentConfigurationFieldItem != null)
            {
                configuratins.Remove(currentConfigurationFieldItem);
            }
            configuratins.Add(configurationFieldItem);
        }

        private static void AddOrUpdateField(IConfigurationItem configuration, IFieldCollection fieldCollection, Mock<IField> mockField)
        {
            var referenceName = mockField.Object.ReferenceName;

            if (fieldCollection.Contains(referenceName))
            {
                var current = fieldCollection.First(x => x.ReferenceName == referenceName);
                fieldCollection.Remove(current);
            }
            var config = configuration.FieldConfigurations.First(f => f.ReferenceFieldName == referenceName);
            mockField.SetupGet(f => f.Configuration).Returns(config);
            fieldCollection.Add(mockField.Object);
        }

        private static Mock<IWorkItem> CreateWorkItem(IConfigurationItem configuration, int id, string propertyPrefix)
        {
            var workItem = new Mock<IWorkItem> { DefaultValue = DefaultValue.Mock };
            workItem.SetupAllProperties();

            var fieldCollection = new FieldCollection();

            foreach (var fieldConfiguration in configuration.FieldConfigurations)
            {
                var field = new Mock<IField> { DefaultValue = DefaultValue.Mock };

                var referenceName = fieldConfiguration.ReferenceFieldName;
                field.SetupAllProperties();

                field.SetupGet(x => x.ReferenceName).Returns(referenceName);
                field.SetupGet(x => x.IsEditable).Returns(true);
                field.SetupGet(x => x.Configuration).Returns(fieldConfiguration);

                if (referenceName == "System.Rev")
                {
                    field.Object.Value = "1";
                }
                else if (referenceName == "System.Id")
                {
                    field.Object.Value = id.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    field.Object.Value = propertyPrefix + referenceName;
                }

                fieldCollection.Add(field.Object);
            }

            workItem.SetupGet(x => x.Id).Returns(id);
            workItem.SetupGet(x => x.Fields).Returns(fieldCollection);
            workItem.SetupGet(x => x.WorkItemType).Returns(configuration.WorkItemType);
            workItem.SetupGet(x => x.Revision).Returns(1);
            workItem.Setup(x => x.GetFieldRevision(It.IsAny<string>())).Returns(1);
            workItem.Setup(x => x.GetWorkItemByRevision(It.IsAny<int>())).Returns(fieldCollection);
            workItem.SetupGet(x => x.Configuration).Returns(configuration);
            workItem.SetupGet(x => x.IsNew).Returns(workItem.Object.Id == 0);
            workItem.Setup(x => x.IsSameWorkItem(It.IsAny<IWorkItem>())).Returns((IWorkItem wi) => (wi.Id != 0 && wi.Id == workItem.Object.Id) || (wi.Id == 0 && wi == workItem.Object));
            if (workItem.Object.Fields.Contains("System.Title"))
            {
                workItem.SetupGet(x => x.Title).Returns(workItem.Object.Fields["System.Title"].Value);
            }
            else
            {
                workItem.SetupGet(x => x.Title).Returns("Value not set");
            }

            return workItem;
        }

        private static Mock<IWorkItemLinkedItems> CreateWorkItemLinkedItems(IConfigurationItem configuration, int id, string propertyPrefix)
        {
            var workItem = new Mock<IWorkItemLinkedItems> { DefaultValue = DefaultValue.Mock };
            workItem.SetupAllProperties();

            var fieldCollection = new FieldCollection();

            foreach (var fieldConfiguration in configuration.FieldConfigurations)
            {
                var field = new Mock<IField> { DefaultValue = DefaultValue.Mock };

                var referenceName = fieldConfiguration.ReferenceFieldName;
                field.SetupAllProperties();

                field.SetupGet(x => x.ReferenceName).Returns(referenceName);
                field.SetupGet(x => x.IsEditable).Returns(true);
                field.SetupGet(x => x.Configuration).Returns(fieldConfiguration);

                if (referenceName == "System.Rev")
                {
                    field.Object.Value = "1";
                }
                else if (referenceName == "System.Id")
                {
                    field.Object.Value = id.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    field.Object.Value = propertyPrefix + referenceName;
                }

                fieldCollection.Add(field.Object);
            }

            workItem.SetupGet(x => x.Id).Returns(id);
            workItem.SetupGet(x => x.Fields).Returns(fieldCollection);
            workItem.SetupGet(x => x.WorkItemType).Returns(configuration.WorkItemType);
            workItem.SetupGet(x => x.Revision).Returns(1);
            workItem.Setup(x => x.GetFieldRevision(It.IsAny<string>())).Returns(1);
            workItem.Setup(x => x.GetWorkItemByRevision(It.IsAny<int>())).Returns(fieldCollection);
            workItem.SetupGet(x => x.Configuration).Returns(configuration);
            workItem.SetupGet(x => x.IsNew).Returns(workItem.Object.Id == 0);
            workItem.Setup(x => x.IsSameWorkItem(It.IsAny<IWorkItemLinkedItems>())).Returns((IWorkItemLinkedItems wi) => (wi.Id != 0 && wi.Id == workItem.Object.Id) || (wi.Id == 0 && wi == workItem.Object));
            if (workItem.Object.Fields.Contains("System.Title"))
            {
                workItem.SetupGet(x => x.Title).Returns(workItem.Object.Fields["System.Title"].Value);
            }
            else
            {
                workItem.SetupGet(x => x.Title).Returns("Value not set");
            }

            return workItem;
        }
        #endregion

        private sealed class WorkItemCollection : Collection<IWorkItem>, IWorkItemCollection
        {
            public IWorkItem Find(int id)
            {
                return Items.Where(x => x.Id == id).FirstOrDefault();
            }
        }

        private sealed class FieldCollection : Collection<IField>, IFieldCollection
        {
            public IField this[string refName]
            {
                get
                {
                    return Items.FirstOrDefault(x => x.ReferenceName.Equals(refName, StringComparison.OrdinalIgnoreCase));
                }
            }

            public bool Contains(string refName)
            {
                return Items.Any(x => x.ReferenceName.Equals(refName, StringComparison.OrdinalIgnoreCase));
            }
        }

        #endregion Helper members
    }
}
