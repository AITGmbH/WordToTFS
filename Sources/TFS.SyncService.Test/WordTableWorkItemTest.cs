namespace TFS.SyncService.Test.Unit
{
    #region Usings
    using System;
    using System.Linq;
    using AIT.TFS.SyncService.Adapter.Word2007;
    using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
    using Microsoft.Office.Interop.Word;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TFS.Test.Common;
    #endregion

    /// <summary>
    ///This is a test class for WordTableWorkItemTest and is intended
    ///to contain all WordTableWorkItemTest Unit Tests
    ///</summary>
    [TestClass]
    public class WordTableWorkItemTest
    {
        #region Fields
        private Document _testDocument;
        #endregion

        #region Properties

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext { get; set; }
        #endregion

       // private Document _testDocument;

        /// <summary>
        /// Create test document with table in it.
        /// </summary>
        [TestInitialize]
        public void MyTestInitialize()
        {
            TFS.Test.Common.TestCleanup.CloseWordDocumentAndKillOpenWordInstances();
            _testDocument = new Document();
            var range = _testDocument.ActiveWindow.Document.Range();
            _testDocument.Tables.Add(range, 6, 5);
            _testDocument.Tables[1].Cell(1, 1).Range.Text = "Requirement";
            _testDocument.Tables[1].Cell(2, 1).Range.Text = "Title";
            _testDocument.Tables[1].Cell(2, 2).Range.Text = "88,89,90"; // links
            _testDocument.Tables[1].Cell(2, 3).Range.Text = "1000"; // id
            _testDocument.Tables[1].Cell(2, 4).Range.Text = "4";
            _testDocument.Tables[1].Cell(3, 1).Range.Text = "Description";

            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
        }

        /// <summary>
        /// Close document
        /// </summary>
        [TestCleanup]
        public void TestCleanup()
        {
            TFS.Test.Common.TestCleanup.CloseWordDocumentAndKillOpenWordInstances();
        }
        //#endregion

        #region TestMethods

        /// <summary>
        ///A test for Links
        ///</summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableWorkItem_GetLinks_ShouldParseCellContent()
        {
            var target = new WordTableWorkItem(_testDocument.Tables[1], "Requirement", CommonConfiguration.Configuration, CommonConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));
            Assert.AreEqual(3, target.Links[target.Links.Keys.First(x => x.LinkValueType == "Child")].Length);
        }

        /// <summary>
        ///A test for Links
        ///</summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableWorkItem_GetLinks_ShouldParseEmptyCellContent()
        {
            _testDocument.Tables[1].Cell(2, 2).Range.Text = string.Empty;
            var target = new WordTableWorkItem(_testDocument.Tables[1], "Requirement", CommonConfiguration.Configuration, CommonConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));
            Assert.AreEqual(0, target.Links[target.Links.Keys.First(x => x.LinkValueType == "Child")].Length);
        }

        /// <summary>
        ///A test for IsNew
        ///</summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableWorkItem_GetIsNew_ShouldReturnTrueForWorkItemWithoutId()
        {
            _testDocument.Tables[1].Cell(2, 3).Range.Text = string.Empty;
            var target2 = new WordTableWorkItem(_testDocument.Tables[1], "Requirement", CommonConfiguration.Configuration, CommonConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));
            Assert.IsTrue(target2.IsNew);
        }

        /// <summary>
        ///A test for IsNew
        ///</summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableWorkItem_GetIsNew_ShouldReturnFalseForWorkItemWithId()
        {
            var target = new WordTableWorkItem(_testDocument.Tables[1], "Requirement", CommonConfiguration.Configuration, CommonConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));
            Assert.IsFalse(target.IsNew);
        }

        /// <summary>
        ///A test for GetRevision
        ///</summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableWorkItem_GetRevision_ShouldReturnValueOfRevisionField()
        {
            var target = new WordTableWorkItem(_testDocument.Tables[1], "Requirement", CommonConfiguration.Configuration, CommonConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));
            Assert.AreEqual(4, target.GetFieldRevision("System.Rev"));
        }

        /// <summary>
        /// Make sure old links are always overwritten
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableWorkItem_AddLinks_ShouldReplaceLinksIfOverwriteIsTrue()
        {
            var target = new WordTableWorkItem(_testDocument.Tables[1], "Requirement", CommonConfiguration.Configuration, CommonConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));
            var fakeAdapter = new Word2007TableSyncAdapter(_testDocument, CommonConfiguration.Configuration);
            fakeAdapter.Open(null);
            target.AddLinks(fakeAdapter, new[] { 4, 5 }, "Child", true);
            Assert.IsTrue(_testDocument.Tables[1].Cell(2, 2).Range.Text.StartsWith("4,5", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Make sure old links are always overwritten
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableWorkItem_AddLinks_ShouldReplaceLinksIfOverwriteIsFalse()
        {
            var target = new WordTableWorkItem(_testDocument.Tables[1], "Requirement", CommonConfiguration.Configuration, CommonConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));
            var fakeAdapter = new Word2007TableSyncAdapter(_testDocument, CommonConfiguration.Configuration);
            fakeAdapter.Open(null);
            target.AddLinks(fakeAdapter, new[] { 4, 5 }, "Child", false);
            Assert.IsTrue(_testDocument.Tables[1].Cell(2, 2).Range.Text.StartsWith("4,5", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Make sure Bookmarks are added and the containing text of the bookmarks is the same
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableWorkItem_AddBookmarks_DocumentAndRequirementShouldContainBookmark()
        {
            var target = new WordTableWorkItem(_testDocument.Tables[1], "Requirement", CommonConfiguration.Configuration, CommonConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));
            var fakeAdapter = new Word2007TableSyncAdapter(_testDocument, CommonConfiguration.Configuration);
            fakeAdapter.Open(null);

            Assert.IsNotNull(target);
            //Desription should contain a Bookmark
            Assert.AreEqual(_testDocument.Tables[1].Cell(3, 1).Range.Bookmarks.Count, 1);
            Assert.AreEqual(_testDocument.Tables[1].Cell(3, 1).Range.Bookmarks[1].Name, CommonConfiguration.BookmarkName);

            //Title should not contain a Bookmark
            Assert.AreNotEqual(_testDocument.Tables[1].Cell(2, 1).Range.Bookmarks.Count, 1);

            //The First Bookmark in the document should be the one from the configuration
            Assert.AreEqual(_testDocument.Bookmarks[1].Name, CommonConfiguration.BookmarkName);

            //The Value of the the Bookmarkshould be the value of the Description
            Assert.IsTrue(_testDocument.Tables[1].Cell(3, 1).Range.Text.Contains(_testDocument.Bookmarks[1].Range.Text));

        }

        /// <summary>
        /// Test if the Bookmark is inserted correctly, even for static value fields
        /// </summary>
        [TestMethod]
        [Ignore] // ALD: Configuration Variables no longer work.
        public void StaticValueField_AddValue_StaticValueFieldShouldContainBookmark()
        {
            var target = new WordTableWorkItem(_testDocument.Tables[1], "Requirement", CommonConfiguration.ConfigurationWithVariables, CommonConfiguration.ConfigurationWithVariables.GetWorkItemConfiguration("Requirement"));
            Assert.IsNotNull(target);
            //The Value of the field that contains the static value field should equal the text from the variable
            Assert.AreEqual(_testDocument.Tables[1].Cell(8, 1).Range.Bookmarks.Count, 1);
            Assert.AreEqual(_testDocument.Tables[1].Cell(8, 1).Range.Bookmarks[1].Name, CommonConfiguration.StaticValueTextBookMarkName);
            Assert.AreNotEqual(_testDocument.Tables[1].Cell(8, 1).Range.Bookmarks[1].Name, CommonConfiguration.BookmarkName);
        }
        #endregion

    }
}
