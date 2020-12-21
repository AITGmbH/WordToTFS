#region Usings
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using AIT.TFS.SyncService.Adapter.Word2007;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Factory;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TFS.Test.Common;
#endregion

namespace TFS.SyncService.Test.Unit
{
    /// <summary>
    /// This is a test class for Word2007SyncAdapterTest and is intended
    /// to contain all Word2007SyncAdapterTest Unit Tests
    /// </summary>
    [TestClass]
    [DeploymentItem("Configuration", "Configuration")]
    public class Word2007SyncAdapterTest
    {
        #region Fields
        private Document _testDocument;
        private Word2007SyncAdapter _wordAdapter;
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the test context which provides
        /// information about and functionality for the current test run.
        /// </summary>
        public TestContext TestContext { get; set; }
        #endregion

        #region Test initializations
        /// <summary>
        /// Create new document
        /// </summary>
        [TestInitialize]
        public void TestInitialize()
        {
            TFS.Test.Common.TestCleanup.CloseWordDocumentAndKillOpenWordInstances();
            _testDocument = new Document();
            _wordAdapter = CreateWord2007SyncAdapter();
        }

        /// <summary>
        /// Cleanup so we are not left with open word application instances
        /// </summary>
        [TestCleanup]
        public void TestCleanup()
        {
            if (_testDocument != null)
            {
                TFS.Test.Common.TestCleanup.CloseWordDocument(_testDocument);
            }
        }
        #endregion

        #region Internal methods
        internal virtual Word2007SyncAdapter CreateWord2007SyncAdapter()
        {
            return new Word2007TableSyncAdapter(_testDocument, CommonConfiguration.Configuration);
        }

        internal virtual Word2007SyncAdapter CreateWord2007SyncAdapterWithConfiguration(IConfiguration configuration)
        {
            return new Word2007TableSyncAdapter(_testDocument, configuration);
        }
        #endregion

        #region TestMethods

        /// <summary>
        /// A test for CreateNewWorkItem
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordAdapter_CreateNewWorkItem_ShouldInsertRelatedTemplate()
        {
            _wordAdapter.Open(null);
            Assert.AreEqual(0, _testDocument.Tables.Count);

            var t = _wordAdapter.CreateNewWorkItem(CommonConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));

            Assert.AreEqual(1, _testDocument.Tables.Count);
            //Assert.AreEqual("ABCD", testDocument.Tables[1].Cell(1, 2).Range.Text); // check if default value is set
            Assert.IsTrue(t.IsNew);
        }

        /// <summary>
        /// A test for GetSelectedWorkItems
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordAdapter_GetSelectedWorkItems_ShouldReturnSelectedWorkItem()
        {
            Debug.Print("Executing Test ShouldReturnSelectedWorkItem");
            SyncServiceTrace.I("Get Range");
            var range = _testDocument.Range();
            CreateRequirementTable(range, "item1");
            var table2 = CreateRequirementTable(range, "item2");
            CreateRequirementTable(range, "item3");

            SyncServiceTrace.I("Select Table");
            table2.Select();
            SyncServiceTrace.I("Get Selected WorkItem WordAdapter");
            var actual = _wordAdapter.GetSelectedWorkItems().ToList();

            Assert.AreEqual(1, actual.Count());
            Assert.AreEqual("item2", actual.Single().Title);
        }

        /// <summary>
        /// A test for Open
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordAdapter_Open_ShouldOpenAllWorkItems()
        {
            var range = _testDocument.Range();
            CreateRequirementTable(range, "item1");
            CreateRequirementTable(range, "item2");

            _wordAdapter.Open(null);

            Assert.AreEqual(2, _wordAdapter.WorkItems.Count);
        }

        /// <summary>
        /// Make sure that when opening a word adapter, existing headers are applied to work items
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordAdapter_Header_OpenWithHeader_ShouldAddHeaderFieldConfigurationToWorkItems()
        {
            var range = _testDocument.Range();

            Create1HeaderTable(range, "\\User Interface", "\\Iteration 2");
            CreateRequirementTable(range);

            _wordAdapter.Open(null);

            Assert.IsTrue(_wordAdapter.WorkItems.First().Configuration.FieldConfigurations.Any(x => x.ReferenceFieldName == "System.IterationPath"));
            Assert.IsTrue(_wordAdapter.WorkItems.First().Configuration.FieldConfigurations.Any(x => x.ReferenceFieldName == "System.AreaPath"));

            Assert.AreEqual("\\User Interface", _wordAdapter.WorkItems.First().Fields["System.AreaPath"].Value);
            Assert.AreEqual("\\Iteration 2", _wordAdapter.WorkItems.First().Fields["System.IterationPath"].Value);
        }

        /// <summary>
        /// Make sure that when opening a word adapter, existing headers are applied to work items
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordAdapter_Header_OpenWithHeader_ShouldAddHeaderFieldsToFieldsWithHierarchyLevel()
        {
            var range = _testDocument.Range();
            Create1HeaderTable(range, "\\User Interface", "\\Iteration 2");
            CreateRequirementTable(range);

            _wordAdapter.Open(null);

            Assert.IsTrue(_wordAdapter.WorkItems.First().Configuration.FieldConfigurations.Any(x => x.ReferenceFieldName == "System.IterationPath"));
            Assert.IsTrue(_wordAdapter.WorkItems.First().Configuration.FieldConfigurations.Any(x => x.ReferenceFieldName == "System.AreaPath"));
            //_wordAdapter.WorkItems.First().Configuration.FieldConfigurations.Add(IConfigurationFieldItem FieldConfigurations.Add());
            //Assert.IsTrue(_wordAdapter.WorkItems.First().Configuration.FieldConfigurations.Any(x => x.ReferenceFieldName == "HierarchyLevel"));
            Assert.AreEqual("\\User Interface", _wordAdapter.WorkItems.First().Fields["System.AreaPath"].Value);
            Assert.AreEqual("\\Iteration 2", _wordAdapter.WorkItems.First().Fields["System.IterationPath"].Value);
        }

        /// <summary>
        /// Stacking header of different levels should combine and replace configurations
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void HeaderTestStackingCombineAndReplaceDifferentLevels()
        {
            var range = _testDocument.ActiveWindow.Document.Range();
            _testDocument.Tables.Add(range, 4, 1);
            _testDocument.Tables[1].Cell(1, 1).Range.Text = "1Header";
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            _testDocument.Tables.Add(range, 4, 1);
            _testDocument.Tables[2].Cell(1, 1).Range.Text = "2Header";
            range = _testDocument.ActiveWindow.Document.Range();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            _testDocument.Tables.Add(range, 5, 5);
            _testDocument.Tables[3].Cell(1, 1).Range.Text = "Requirement";

            // work item value is empty =>  read from header
            var target = CreateWord2007SyncAdapter();

            Assert.IsTrue(target.Open(null));

            // test is only useful if prerequisites are met
            Assert.AreNotEqual(Direction.SetInNewTfsWorkItem, CommonConfiguration.Header1.Fields.First(x => x.Name == "System.AreaPath").Direction);
            Assert.AreEqual(Direction.SetInNewTfsWorkItem, CommonConfiguration.Header2.Fields.First(x => x.Name == "System.AreaPath").Direction);

            Assert.AreEqual(Direction.SetInNewTfsWorkItem, target.WorkItems.First().Fields["System.AreaPath"].Configuration.Direction);
            Assert.IsTrue(target.WorkItems.First().Fields.Contains("System.GibtsNicht"));
            Assert.IsTrue(target.WorkItems.First().Fields.Contains("System.IterationPath")); // inherited from level 1 header
        }

        /// <summary>
        /// Make sure that when refreshing from TFS, Header fields are not written to.
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordAdapter_Header_SetWorkItemHeaderField_ShouldNotWriteToHeaderFields()
        {
            var range = _testDocument.Range();
            Create1HeaderTable(range, "\\User Interface");
            CreateRequirementTable(range);

            _wordAdapter.Open(null);
            var workItemField = _wordAdapter.WorkItems.First().Fields["System.AreaPath"];
            workItemField.Value = "bongo";

            Assert.AreEqual("\\User Interface", workItemField.Value);
        }

        /// <summary>
        /// If Fields are defined in headers and work items, a split field is created in the word adapter and the configuration
        /// it presented is that of the work item definition
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordAdapter_Header_OpenItemWithSameFieldConfigurationAsHeader_ShouldCreateSplitField()
        {
            var range = _testDocument.Range();
            Create1HeaderTable(range);
            CreateRequirementTable(range);

            _wordAdapter.Open(null);

            Assert.AreEqual(Direction.OtherToTfs, CommonConfiguration.Configuration.GetConfigurationItems().First(x => x.WorkItemTypeMapping == "Requirement").FieldConfigurations.First(x => x.ReferenceFieldName == "System.Title").Direction);
            Assert.AreEqual(Direction.SetInNewTfsWorkItem, CommonConfiguration.Configuration.Headers.First(x => x.WorkItemTypeMapping == "1Header").FieldConfigurations.First(x => x.ReferenceFieldName == "System.Title").Direction);

            // make sure header parsing noticed both configs are present
            Assert.IsTrue(_wordAdapter.WorkItems.First().Fields["System.Title"] is AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects.SplitField);
        }

        /// <summary>
        /// When a header field is set to <c>SetInNewTfs</c>, its value should only be taken for new work items.
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordAdapter_Header_FieldWithSetInNewTfsWorkItem_ShouldSetValueOnlyForNewWorkItems()
        {
            var range = _testDocument.ActiveWindow.Document.Range();
            Create1HeaderTable(range, string.Empty, string.Empty, "HeaderTitle");
            var workItemTable = CreateRequirementTable(range);

            // work item value is empty => read from header
            _wordAdapter.Open(null);
            Assert.AreEqual("HeaderTitle", _wordAdapter.WorkItems.First().Fields["System.Title"].Value);

            // overwrite title in work item for new wi => read from wi
            workItemTable.Cell(2, 1).Range.Text = "Title";
            var target = CreateWord2007SyncAdapter();
            target.Open(null);
            Assert.AreEqual("Title", target.WorkItems.First().Fields["System.Title"].Value);

            // now give the requirement an id (isNew = false), but delete its title => still read from wi
            workItemTable.Cell(2, 3).Range.Text = "1";
            workItemTable.Cell(2, 1).Range.Text = string.Empty;
            target = CreateWord2007SyncAdapter();
            target.Open(null);
            Assert.AreEqual(string.Empty, target.WorkItems.First().Fields["System.Title"].Value);
        }

        /// <summary>
        /// When a header field is set to <c>SetInNewTfs</c>, its value should only be taken for new work items.
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void HeaderTestWriteSplitFieldWithSetInNewTfsWorkItem()
        {
            var range = _testDocument.ActiveWindow.Document.Range();
            _testDocument.Tables.Add(range, 4, 1);
            _testDocument.Tables[1].Cell(1, 1).Range.Text = "1Header";
            _testDocument.Tables[1].Cell(4, 1).Range.Text = "HeaderTitle";

            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            _testDocument.Tables.Add(range, 5, 5);
            _testDocument.Tables[2].Cell(1, 1).Range.Text = "Requirement";
            _testDocument.Tables[2].Select();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            // work item value is empty =>  read from header
            var target = CreateWord2007SyncAdapter();
            Assert.IsTrue(target.Open(null));

            target.WorkItems.First().Fields["System.Title"].Value = "new title";

            Assert.AreEqual("HeaderTitle", _testDocument.Tables[1].Cell(4, 1).Range.Text.TrimEnd(new[] { (char)7, (char)13 }));
            Assert.AreEqual("new title", _testDocument.Tables[2].Cell(2, 1).Range.Text.TrimEnd(new[] { (char)7, (char)13 }));
        }

        /// <summary>
        /// Test for Bug # 17650
        /// When using the enhanced Document Structure the field contain subtype "HierarchyLevel"
        /// The Word Adapter cant handle it and an error is produced
        /// 
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void BehaviorOfWordSyncAdapterWithUnknownFieldValues()
        {
            const string hierarachyLevel0TemplateName = "HierarchieTemplateLevel0";
            const string hierarachyLevel1TemplateName = "HierarchieTemplateLevel1";

            //Create Configuration
            IConfiguration conf = CommonConfiguration.Configuration;
            var featureBaseConfiguration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title").GetConfigurationItems().Single();
            var featureHierarchieLevel0Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel0TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "0").GetConfigurationItems().Single();
            var featureHierarchieLevel1Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel1TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "1").GetConfigurationItems().Single();

            conf.GetConfigurationItems().Add(featureBaseConfiguration);
            conf.GetConfigurationItems().Add(featureHierarchieLevel0Configuration);
            conf.GetConfigurationItems().Add(featureHierarchieLevel1Configuration);

            var range = _testDocument.ActiveWindow.Document.Range();
            _testDocument.Tables.Add(range, 4, 1);
            _testDocument.Tables[1].Cell(1, 1).Range.Text = "1Header";
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            _testDocument.Tables.Add(range, 4, 1);
            _testDocument.Tables[2].Cell(1, 1).Range.Text = "2Header";
            range = _testDocument.ActiveWindow.Document.Range();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            _testDocument.Tables.Add(range, 5, 5);
            _testDocument.Tables[3].Cell(1, 1).Range.Text = "Requirement";


            // Open the adapter with the configuration that contains the "HierarchyLevel"
            var target = CreateWord2007SyncAdapterWithConfiguration(conf);

            //Open the adapter
            Assert.IsTrue(target.Open(null));

            // test is only useful if prerequisites are met
            Assert.AreNotEqual(Direction.SetInNewTfsWorkItem, CommonConfiguration.Header1.Fields.First(x => x.Name == "System.AreaPath").Direction);
            Assert.AreEqual(Direction.SetInNewTfsWorkItem, CommonConfiguration.Header2.Fields.First(x => x.Name == "System.AreaPath").Direction);

            Assert.AreEqual(Direction.SetInNewTfsWorkItem, target.WorkItems.First().Fields["System.AreaPath"].Configuration.Direction);
            Assert.IsTrue(target.WorkItems.First().Fields.Contains("System.GibtsNicht"));
            Assert.IsTrue(target.WorkItems.First().Fields.Contains("System.IterationPath")); // inherited from level 1 header
        }

        /// <summary>
        /// Of two header of the same level, the second one replaces all definitions of the first one in its range
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void HeaderTestStackingReplaceHeaderOfSameOrHigherLevel()
        {
            var range = _testDocument.ActiveWindow.Document.Range();
            _testDocument.Tables.Add(range, 4, 1);
            _testDocument.Tables[1].Cell(1, 1).Range.Text = "1Header";
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            _testDocument.Tables.Add(range, 4, 1);
            _testDocument.Tables[2].Cell(1, 1).Range.Text = "2Header";
            range = _testDocument.ActiveWindow.Document.Range();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            _testDocument.Tables.Add(range, 4, 1);
            _testDocument.Tables[3].Cell(1, 1).Range.Text = "1HeaderEnd";
            range = _testDocument.ActiveWindow.Document.Range();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            _testDocument.Tables.Add(range, 5, 5);
            _testDocument.Tables[4].Cell(1, 1).Range.Text = "Requirement";

            var target = CreateWord2007SyncAdapter();
            Assert.IsTrue(target.Open(null));

            // test is only useful if prerequisites are met
            Assert.IsTrue(CommonConfiguration.Header2.Fields.Any(x => x.Name == "System.GibtsNicht"));
            Assert.IsTrue(CommonConfiguration.Header1.Fields.Any(x => x.Name == "System.AreaPath"));
            Assert.IsTrue(CommonConfiguration.Header1.Fields.Any(x => x.Name == "System.IterationPath"));
            Assert.IsTrue(CommonConfiguration.Header1.Fields.Any(x => x.Name == "System.Title"));

            // make sure none of the header1 exclusive fields is applied
            Assert.IsFalse(target.WorkItems.First().Fields.Contains("System.GibtsNicht"));
            Assert.AreNotEqual(Direction.SetInNewTfsWorkItem, target.WorkItems.First().Fields["System.Title"].Configuration.Direction);
        }

        /// <summary>
        /// Check if tables are checked for having exactly the same structure as in the related template
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        [DeploymentItem("Configuration", "Configuration")]
        public void WordSyncAdapter_Validate_ShouldReturnErrorIfTableIsDifferentToRelatedTemplate()
        {
            var issueConfig = CommonConfiguration.Configuration.GetConfigurationItems().Where(x => x.WorkItemType == "Issue").First();
            _testDocument.Range().InsertFile(Path.Combine(Environment.CurrentDirectory, issueConfig.RelatedTemplateFile));
            _testDocument.Tables[1].Rows.Add();

            var target = CreateWord2007SyncAdapter();
            target.Open(null);

            Assert.AreEqual(1, target.ValidateWorkItems().Count);
        }

        /// <summary>
        /// Check if tables are checked for having exactly the same structure as in the related template
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        [DeploymentItem("Configuration", "Configuration")]
        public void WordSyncAdapter_Validate_ShouldNotReturnError()
        {
            var issueConfig = CommonConfiguration.Configuration.GetConfigurationItems().Where(x => x.WorkItemType == "Issue").First();
            _testDocument.Range().InsertFile(Path.Combine(Environment.CurrentDirectory, issueConfig.RelatedTemplateFile));

            var target = CreateWord2007SyncAdapter();
            target.Open(null);

            Assert.AreEqual(0, target.ValidateWorkItems().Count);
        }
        #endregion

        #region Private methods
        private Table CreateRequirementTable(Range range, string title = "")
        {
            SyncServiceTrace.I("Create Requirement Table");
            var table = _testDocument.Tables.Add(range, 5, 4);
            table.Cell(1, 1).Range.Text = "Requirement";
            table.Cell(1, 4).Range.Text = string.Empty; // State
            table.Cell(2, 1).Range.Text = title;
            table.Cell(2, 3).Range.Text = string.Empty; // id
            table.Cell(2, 4).Range.Text = string.Empty; // revision
            table.Cell(3, 1).Range.Text = string.Empty; // description
            table.Cell(4, 1).Range.Text = string.Empty; // AreaPath
            table.Cell(5, 1).Range.Text = string.Empty; // IterationPath

            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            return table;
        }

        private void Create1HeaderTable(Range range, string areaPath = "", string iteration = "", string title = "")
        {
            var table = _testDocument.Tables.Add(range, 4, 1);
            table.Cell(1, 1).Range.Text = "1Header";
            table.Cell(2, 1).Range.Text = areaPath;
            table.Cell(3, 1).Range.Text = iteration;
            table.Cell(4, 1).Range.Text = title;

            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
        }
        #endregion

    }
}
