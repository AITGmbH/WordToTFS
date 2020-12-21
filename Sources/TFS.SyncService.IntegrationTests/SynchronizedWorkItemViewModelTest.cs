#region Usings
using System;
using System.Linq;
using AIT.TFS.SyncService.Adapter.TFS2012;
using AIT.TFS.SyncService.Adapter.Word2007;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Model;
using AIT.TFS.SyncService.Service.Configuration;
using Microsoft.Office.Interop.Word;
using Microsoft.TeamFoundation.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using TFS.Test.Common;
// ReSharper disable InconsistentNaming
#endregion

namespace TFS.SyncService.Test.Integration
{
    /// <summary>
    /// Test the class SynchronizedWorkItemViewModel
    /// </summary>
    [TestClass]
    public class SynchronizedWorkItemViewModelTest
    {
        #region Private fields
        private SynchronizedWorkItemViewModel _synchronizedWorkItemViewModel;
        private IWorkItemSyncAdapter _tfsAdapter;
        private IWorkItemSyncAdapter _wordAdapter;
        private Document _document;
        private Mock<ISyncServiceDocumentModel> _documentModel;
        // ReSharper disable once NotAccessedField.Local
        //This is needed to make sure that the test server is authorized to connect to the tfs
        private static TfsTeamProjectCollection projectCollection;

        private static TestContext _testContext;

        #endregion

        #region Test Initializations and Cleanup

        /// <summary>
        /// Make sure all services are registered.
        /// </summary>
        [ClassInitialize]
        public static void MyClassInitialize(TestContext testContext)
        {
            _testContext = testContext;
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);

            // ReSharper disable once CSharpWarnings::CS0618
            projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(CommonConfiguration.TfsTestServerConfiguration(_testContext).TeamProjectCollectionUrl));

            AIT.TFS.SyncService.Service.AssemblyInit.Instance.Init();
            AIT.TFS.SyncService.Adapter.TFS2012.AssemblyInit.Instance.Init();
            AIT.TFS.SyncService.Adapter.Word2007.AssemblyInit.Instance.Init();
       }

        /// <summary>
        /// Initialize tests
        /// </summary>
        [TestInitialize]
        public void TestInitialize()
        {
            TestCleanup.CloseWordDocumentAndKillOpenWordInstances();
            _document = new Document();
        }

        #endregion Test Initializations and Cleanup


        #region TestMethods
        /// <summary>
        /// If the work item does not exist in the document, it should be "Not Imported"
        /// This test need a timeout as it opens new word document ins the background. this can lead to problems with the COM and the testing
        /// </summary>
        [TestMethod, Timeout(30000)]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void SynchronizedWorkItemViewModel_SynchronizationState_ShouldBeNotImported()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");
            OpenAdapters(CommonConfiguration.TfsTestServerConfiguration(_testContext), config);

            _synchronizedWorkItemViewModel = new SynchronizedWorkItemViewModel(_tfsAdapter.WorkItems.First(), null, _wordAdapter, _documentModel.Object, null);
            _synchronizedWorkItemViewModel.Refresh();

            Assert.AreEqual(SynchronizationState.NotImported, _synchronizedWorkItemViewModel.SynchronizationState);
        }

        /// <summary>
        /// If the work item exists and all fields are the same as on the server, it should be "Up To Date"
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration", "Configuration")]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void SynchronizedWorkItemViewModel_SynchronizationState_ShouldBeUpToDate()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");
            OpenAdapters(CommonConfiguration.TfsTestServerConfiguration(_testContext), config);

            _wordAdapter.CreateNewWorkItem(config.GetConfigurationItems().First());
            _wordAdapter.WorkItems.First().Fields["System.Title"].Value = _tfsAdapter.WorkItems.First().Fields["System.Title"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Id"].Value = _tfsAdapter.WorkItems.First().Fields["System.Id"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Rev"].Value = _tfsAdapter.WorkItems.First().Fields["System.Rev"].Value;
            _wordAdapter = new Word2007TableSyncAdapter(_document, config);
            _wordAdapter.Open(null);

             _synchronizedWorkItemViewModel = new SynchronizedWorkItemViewModel(_tfsAdapter.WorkItems.First(), null, _wordAdapter, _documentModel.Object, null);
            _synchronizedWorkItemViewModel.Refresh();
            Assert.AreEqual(SynchronizationState.UpToDate, _synchronizedWorkItemViewModel.SynchronizationState);
        }

        /// <summary>
        /// If the work item exists and all fields are the same as on the server, it should be "Up To Date"
        /// This is an extended test for the case, that stackRank is set to true. This value should be ignored by the syncservice and the WorkItems should be UpToDate
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration", "Configuration")]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void SynchronizedWorkItemViewModel_SynchronizationState_ShouldBeUpToDate_WithStackRankSetToTrue()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");
            OpenAdapters(CommonConfiguration.TfsTestServerConfiguration(_testContext), config);
            config.UseStackRank = true;

            //add the stack rank configuration and a value
            _tfsAdapter.WorkItems.First().Configuration.FieldConfigurations.Add(new ConfigurationFieldItem("Microsoft.VSTS.Common.StackRank", string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 0, 0, string.Empty, false, HandleAsDocumentType.All, null, string.Empty, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));
            _tfsAdapter.WorkItems.First().Fields["Microsoft.VSTS.Common.StackRank"].Value = "3000";

            _wordAdapter.CreateNewWorkItem(config.GetConfigurationItems().First());
            _wordAdapter.WorkItems.First().Fields["System.Title"].Value = _tfsAdapter.WorkItems.First().Fields["System.Title"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Id"].Value = _tfsAdapter.WorkItems.First().Fields["System.Id"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Rev"].Value = _tfsAdapter.WorkItems.First().Fields["System.Rev"].Value;
            _wordAdapter.WorkItems.First().Configuration.FieldConfigurations.Add(new ConfigurationFieldItem("Microsoft.VSTS.Common.StackRank", string.Empty, FieldValueType.PlainText, Direction.OtherToTfs, 0, 0, string.Empty, false, HandleAsDocumentType.All, null, string.Empty, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null));

            _wordAdapter = new Word2007TableSyncAdapter(_document, config);
            _wordAdapter.Open(null);

            _synchronizedWorkItemViewModel = new SynchronizedWorkItemViewModel(_tfsAdapter.WorkItems.First(), null, _wordAdapter, _documentModel.Object, null);
            _synchronizedWorkItemViewModel.Refresh();
            Assert.AreEqual(SynchronizationState.UpToDate, _synchronizedWorkItemViewModel.SynchronizationState);
        }

        /// <summary>
        /// If the work item exists in the document, but the server has at least one field
        /// that has changed, this work item should be "Outdated"
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration", "Configuration")]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void SynchronizedWorkItemViewModel_SynchronizationState_ShouldBeOutdated()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");
            OpenAdapters(CommonConfiguration.TfsTestServerConfiguration(_testContext), config);

            var tfsItem = _tfsAdapter.WorkItems.First();
            var oldFields = tfsItem.GetWorkItemByRevision(tfsItem.Revision - 1);

            _wordAdapter.CreateNewWorkItem(config.GetConfigurationItems().First());
            _wordAdapter.WorkItems.First().Fields["System.Title"].Value = oldFields["System.Title"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Id"].Value = oldFields["System.Id"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Rev"].Value = oldFields["System.Rev"].Value;

            _wordAdapter = new Word2007TableSyncAdapter(_document, config);
            _wordAdapter.Open(null);

            _synchronizedWorkItemViewModel = new SynchronizedWorkItemViewModel(_tfsAdapter.WorkItems.First(), null, _wordAdapter, _documentModel.Object, null);
            _synchronizedWorkItemViewModel.Refresh();
            Assert.AreEqual(SynchronizationState.Outdated, _synchronizedWorkItemViewModel.SynchronizationState);
        }

        /// <summary>
        /// If the work item exists in the document and some fields have changed, but these
        /// changes will never be published and the server is of the same revision
        /// the local work item should be "Differing"
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration", "Configuration")]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void SynchronizedWorkItemViewModel_SynchronizationState_ShouldBeDiffering()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.TfsToOther, FieldValueType.PlainText, "System.Title");
            OpenAdapters(CommonConfiguration.TfsTestServerConfiguration(_testContext), config);

            var tfsItem = _tfsAdapter.WorkItems.First();

            _wordAdapter.CreateNewWorkItem(config.GetConfigurationItems().First());
            _wordAdapter.WorkItems.First().Fields["System.Title"].Value = "changedTitle";
            _wordAdapter.WorkItems.First().Fields["System.Id"].Value = tfsItem.Fields["System.Id"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Rev"].Value = tfsItem.Fields["System.Rev"].Value;
            _wordAdapter = new Word2007TableSyncAdapter(_document, config);
            _wordAdapter.Open(null);

            _synchronizedWorkItemViewModel = new SynchronizedWorkItemViewModel(_tfsAdapter.WorkItems.First(), null, _wordAdapter, _documentModel.Object, null);
            _synchronizedWorkItemViewModel.Refresh();
            Assert.AreEqual(SynchronizationState.Differing, _synchronizedWorkItemViewModel.SynchronizationState);
        }

        /// <summary>
        /// If the work item exists in the document and some fields have changed it should be "Dirty"
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration", "Configuration")]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void SynchronizedWorkItemViewModel_SynchronizationState_ShouldBeDirty()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");
            OpenAdapters(CommonConfiguration.TfsTestServerConfiguration(_testContext), config);

            var tfsItem = _tfsAdapter.WorkItems.First();

            _wordAdapter.CreateNewWorkItem(config.GetConfigurationItems().First());
            _wordAdapter.WorkItems.First().Fields["System.Title"].Value = "TestMichDoch";
            _wordAdapter.WorkItems.First().Fields["System.Id"].Value = tfsItem.Fields["System.Id"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Rev"].Value = tfsItem.Fields["System.Rev"].Value;

            _wordAdapter = new Word2007TableSyncAdapter(_document, config);
            _wordAdapter.Open(null);
            _synchronizedWorkItemViewModel = new SynchronizedWorkItemViewModel(_tfsAdapter.WorkItems.First(), null, _wordAdapter, _documentModel.Object, null);

            _synchronizedWorkItemViewModel.Refresh();
            Assert.AreEqual(SynchronizationState.Dirty, _synchronizedWorkItemViewModel.SynchronizationState);
        }

        /// <summary>
        /// If the work item exists in the document and both, the Word and the TFS version have changed
        /// different field values and can still be merged, the state should be "Diverged without conflicts"
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration", "Configuration")]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void SynchronizedWorkItemViewModel_SynchronizationState_ShouldBeDivergedWithConflicts()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");
            config.GetConfigurationItems().First().FieldConfigurations.Add(new ConfigurationFieldItem("System.AreaPath", "", FieldValueType.PlainText, Direction.OtherToTfs, 5, 1, "", false, HandleAsDocumentType.OleOnDemand, null, null, null, ShapeOnlyWorkaroundMode.AddSpace,null,null,null));

            OpenAdapters(CommonConfiguration.TfsTestServerConfiguration(_testContext), config);
            var tfsItem = _tfsAdapter.WorkItems.First();
            var oldFields = tfsItem.GetWorkItemByRevision(tfsItem.Revision - 1);

            Assert.AreNotEqual(tfsItem.Fields["System.AreaPath"].Value, oldFields["System.AreaPath"].Value);

            _wordAdapter.CreateNewWorkItem(config.GetConfigurationItems().First());
            _wordAdapter.WorkItems.First().Fields["System.Title"].Value = "Test";
            _wordAdapter.WorkItems.First().Fields["System.Id"].Value = oldFields["System.Id"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Rev"].Value = oldFields["System.Rev"].Value;
            _wordAdapter.WorkItems.First().Fields["System.AreaPath"].Value = oldFields["System.AreaPath"].Value;
            _wordAdapter = new Word2007TableSyncAdapter(_document, config);
            _wordAdapter.Open(null);

            _synchronizedWorkItemViewModel = new SynchronizedWorkItemViewModel(_tfsAdapter.WorkItems.First(), null, _wordAdapter, _documentModel.Object, null);
            _synchronizedWorkItemViewModel.Refresh();
            Assert.AreEqual(SynchronizationState.DivergedWithConflicts, _synchronizedWorkItemViewModel.SynchronizationState);
        }

        /// <summary>
        /// If the work item exists in the document and both, the Word and the TFS version
        /// have changed the same field, the state should be "Diverged with conflicts".
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration", "Configuration")]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void SynchronizedWorkItemViewModel_SynchronizationState_ShouldBeDivergedWithoutConflicts()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");

            OpenAdapters(CommonConfiguration.TfsTestServerConfiguration(_testContext), config);
            var tfsItem = _tfsAdapter.WorkItems.First();
            var oldFields = tfsItem.GetWorkItemByRevision(tfsItem.Revision - 1);

            _wordAdapter.CreateNewWorkItem(config.GetConfigurationItems().First());
            _wordAdapter.WorkItems.First().Fields["System.Title"].Value = "Was anderes";
            _wordAdapter.WorkItems.First().Fields["System.Id"].Value = oldFields["System.Id"].Value;
            _wordAdapter.WorkItems.First().Fields["System.Rev"].Value = "1";
            _wordAdapter = new Word2007TableSyncAdapter(_document, config);
            _wordAdapter.Open(null);

            _synchronizedWorkItemViewModel = new SynchronizedWorkItemViewModel(_tfsAdapter.WorkItems.First(), null, _wordAdapter, _documentModel.Object, null);
            _synchronizedWorkItemViewModel.Refresh();
            Assert.AreEqual(SynchronizationState.DivergedWithoutConflicts, _synchronizedWorkItemViewModel.SynchronizationState);
        }

        /// <summary>
        /// If the work item exists in the document but not on the server,
        /// its state should be "New".
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration", "Configuration")]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void SynchronizedWorkItemViewModel_SynchronizationState_ShouldBeNew()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");

            OpenAdapters(CommonConfiguration.TfsTestServerConfiguration(_testContext), config);
            _wordAdapter.CreateNewWorkItem(config.GetConfigurationItems().First());
            _wordAdapter.WorkItems.First().Fields["System.Title"].Value = "Was anderes";
            _wordAdapter = new Word2007TableSyncAdapter(_document, config);
            _wordAdapter.Open(null);

            _synchronizedWorkItemViewModel = new SynchronizedWorkItemViewModel(null, _wordAdapter.WorkItems.First(), _wordAdapter, _documentModel.Object, null);
            _synchronizedWorkItemViewModel.Refresh();
            Assert.AreEqual(SynchronizationState.New, _synchronizedWorkItemViewModel.SynchronizationState);
        }

        #endregion

        #region TestHelpers

        private void OpenAdapters(ServerConfiguration serverConfig, IConfiguration configuration)
        {
            _tfsAdapter = new Tfs2012SyncAdapter(serverConfig.TeamProjectCollectionUrl, serverConfig.TeamProjectName, null, configuration);
            _tfsAdapter.Open(new[]
                                {
                                    serverConfig.SynchronizedWorkItemId
                                });

            _documentModel = new Mock<ISyncServiceDocumentModel>();
            _documentModel.SetupGet(x => x.Configuration).Returns(configuration);
            _documentModel.SetupGet(x => x.TfsProject).Returns(serverConfig.TeamProjectName);
            _documentModel.SetupGet(x => x.TfsServer).Returns(serverConfig.TeamProjectCollectionUrl);

            _wordAdapter = new Word2007TableSyncAdapter(_document, configuration);
            _wordAdapter.Open(null);
        }

        #endregion TestHelpers
    }
}
