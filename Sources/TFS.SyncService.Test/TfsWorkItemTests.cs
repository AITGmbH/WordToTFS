// ReSharper disable InconsistentNaming
namespace TFS.SyncService.Test.Unit
{
    using System;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using AIT.TFS.SyncService.Adapter.TFS2012;
    using AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects;
    using AIT.TFS.SyncService.Contracts.Configuration;
    using AIT.TFS.SyncService.Contracts.WorkItemObjects;
    using Microsoft.TeamFoundation.Client;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TFS.Test.Common;

    /// <summary>
    /// This is a test class that contains all tests for <see cref="TfsWorkItem"/>.
    /// </summary>
    [TestClass]
    [DeploymentItem(@"..\packages\Microsoft.TeamFoundationServer.ExtendedClient.16.153.0\lib\native\x86\Microsoft.WITDataStore32.dll")]
    [DeploymentItem(@"..\packages\Microsoft.TeamFoundationServer.ExtendedClient.16.153.0\lib\native\x86\Microsoft.Microsoft.WITDataStore64.dll")]
    public class TfsWorkItemTests
    {
        #region Context of each test

        /// <summary>
        /// Gets or sets the test context which provides
        /// information about and functionality for the current test run.
        /// </summary>
        public TestContext TestContext { get; set; }

        private ServerConfiguration _testServerConfiguration;
        private TfsWorkItem _requirement;
        private TfsWorkItem _issue;
        private Tfs2012SyncAdapter _adapter;
        private static TestContext _testContext;

        #endregion Context of each test

        #region Test initializations

        /// <summary>
        /// Make sure all services are registered.
        /// </summary>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            _testContext = testContext;
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);
            AIT.TFS.SyncService.Service.AssemblyInit.Instance.Init();
        }

        /// <summary>
        /// Create adapter and get work items relevant for these tests.
        /// </summary>
        [TestInitialize]
        [TestCategory("ConnectionNeeded")]
        public void TestInitialize()
        {
            _testServerConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            // ReSharper disable once UnusedVariable
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(_testServerConfiguration.TeamProjectCollectionUrl));
            CreateAdapterForTeamProject();
            GetWorkItemsRelevantForTheseTests();
        }

        private void CreateAdapterForTeamProject()
        {
            _adapter = new Tfs2012SyncAdapter(_testServerConfiguration.TeamProjectCollectionUrl, _testServerConfiguration.TeamProjectName, null, _testServerConfiguration.Configuration);
            var resultOfAdapterOpening = _adapter.Open(new[] { CommonConfiguration.RequirementId, CommonConfiguration.IssueId });
            InvalidateTestIfNotTrue(resultOfAdapterOpening, "Adapter.Open() did not succeed.");
        }

        private void GetWorkItemsRelevantForTheseTests()
        {
            _requirement = (TfsWorkItem)_adapter.WorkItems.Find(CommonConfiguration.RequirementId);
            _issue = (TfsWorkItem)_adapter.WorkItems.Find(CommonConfiguration.IssueId);
        }

        #endregion Test initializations

        #region TestMethods

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksWithEmptyIdListAndOverwriteFlag ShouldDeleteAllLinksOfTheGivenLinkValueType.
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinksWithEmptyIdListAndOverwriteFlag_ShouldDeleteAllLinksOfTheGivenLinkValueType()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);
            _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId }, linkConfiguration.LinkValueType, true);
            InvalidateTestIfNotEqual(1, GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Length, "workItem.AddLinks() did not create all links successfully.");

            // act
            var addLinksReturnValue = _requirement.AddLinks(_adapter, new int[] { }, linkConfiguration.LinkValueType, true);

            // assert
            Assert.IsTrue(addLinksReturnValue, "Unexpected return value of WorkItem.AddLinks().");
            Assert.AreEqual(0, GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Length, "Count of links does not match.");
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksBetweenWorkItemsOfSameTeamProject ShouldAddLink.
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinksBetweenWorkItemsOfSameTeamProject_ShouldAddLink()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            var addLinksReturnValue = _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId }, linkConfiguration.LinkValueType, true);

            // assert
            Assert.IsTrue(addLinksReturnValue, "Unexpected return value of WorkItem.AddLinks().");
            Assert.IsTrue(GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Contains(CommonConfiguration.IssueId), "Work item id was not found in links.");
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksBetweenWorkItemsOfDifferentTeamProjects ShouldThrowException.
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinksBetweenWorkItemsOfDifferentTeamProjects_ShouldAddLink()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            var addLinksReturnValue = _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IdOfWorkItemInOtherTeamProject }, linkConfiguration.LinkValueType, true);

            // assert
            Assert.IsTrue(addLinksReturnValue, "Unexpected return value of WorkItem.AddLinks().");
            Assert.IsTrue(GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Contains(CommonConfiguration.IdOfWorkItemInOtherTeamProject), "Work item id was not found in links.");
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksForAWorkItemTypeDefinedAsASingleLinkedWorkItemType ShouldAddLink
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinksForAWorkItemTypeDefinedAsASingleLinkedWorkItemType_ShouldAddLink()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            var addLinksReturnValue = _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId }, linkConfiguration.LinkValueType, "Issue", false);

            // assert
            Assert.IsTrue(addLinksReturnValue, "Unexpected return value of WorkItem.AddLinks().");
            Assert.IsTrue(GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Contains(CommonConfiguration.IssueId), "Work item id was not found in links.");
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksForAWorkItemTypeNotDefinedInLinkedWorkItemTypes ShouldThrowException.
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        [ExpectedException(typeof(ArgumentException))]
        public void TfsWorkItem_AddLinksForAWorkItemTypeNotDefinedInLinkedWorkItemTypes_ShouldThrowException()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId }, linkConfiguration.LinkValueType, "Bug", false);

            // assert exception
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksForAWorkItemTypeDefinedInAListOfLinkedWorkItemTypes ShouldAddLink
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinksForAWorkItemTypeDefinedInAListOfLinkedWorkItemTypes_ShouldAddLink()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            var addLinksReturnValue = _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId }, linkConfiguration.LinkValueType, "Issue,Requirement", false);

            // assert
            Assert.IsTrue(addLinksReturnValue, "Unexpected return value of WorkItem.AddLinks().");
            Assert.IsTrue(GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Contains(CommonConfiguration.IssueId), "Work item id was not found in links.");
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksForAWorkItemTypeDefinedInAListOfLinkedWorkItemTypesWithWhitespace ShouldAddLink
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinksForAWorkItemTypeDefinedInAListOfLinkedWorkItemTypesWithWhitespace_ShouldAddLink()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            var addLinksReturnValue = _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId }, linkConfiguration.LinkValueType, "Requirement, Issue", false);

            // assert
            Assert.IsTrue(addLinksReturnValue, "Unexpected return value of WorkItem.AddLinks().");
            Assert.IsTrue(GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Contains(CommonConfiguration.IssueId), "Work item id was not found in links.");
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksForAWorkItemTypeWithEmptyLinkedWorkItemTypes ShouldAddLink
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinksForAWorkItemTypeWithEmptyLinkedWorkItemTypes_ShouldAddLink()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            var addLinksReturnValue = _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId }, linkConfiguration.LinkValueType, string.Empty, false);

            // assert
            Assert.IsTrue(addLinksReturnValue, "Unexpected return value of WorkItem.AddLinks().");
            Assert.IsTrue(GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Contains(CommonConfiguration.IssueId), "Work item id was not found in links.");
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksForAWorkItemTypeWithNullAsLinkedWorkItemTypes ShouldAddLink
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinksForAWorkItemTypeWithNullAsLinkedWorkItemTypes_ShouldAddLink()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            var addLinksReturnValue = _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId }, linkConfiguration.LinkValueType, null, false);

            // assert
            Assert.IsTrue(addLinksReturnValue, "Unexpected return value of WorkItem.AddLinks().");
            Assert.IsTrue(GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Contains(CommonConfiguration.IssueId), "Work item id was not found in links.");
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksForMultipleWorkItemTypesDefinedInAListOfLinkedWorkItemTypes ShouldAddLinks
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinksForMultipleWorkItemTypesDefinedInAListOfLinkedWorkItemTypes_ShouldAddLinks()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            var addLinksReturnValue = _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId, CommonConfiguration.TestCaseId }, linkConfiguration.LinkValueType, "Test Case,Issue", false);

            // assert
            Assert.IsTrue(addLinksReturnValue, "Unexpected return value of WorkItem.AddLinks().");
            Assert.IsTrue(GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Contains(CommonConfiguration.IssueId), "Work item id was not found in links.");
            Assert.IsTrue(GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement).Contains(CommonConfiguration.TestCaseId), "Work item id was not found in links.");
        }

        /// <summary>
        /// Tests <see cref="TfsWorkItem"/> AddLinksForMultipleWorkItemTypesNotFullyDefinedInAListOfLinkedWorkItemTypes ShouldThrowException
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        [ExpectedException(typeof(ArgumentException))]
        public void TfsWorkItem_AddLinksForMultipleWorkItemTypesNotFullyDefinedInAListOfLinkedWorkItemTypes_ShouldThrowException()
        {
            // arrange
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Child);
            RemoveExistingLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, _requirement);

            // act
            _requirement.AddLinks(_adapter, new[] { CommonConfiguration.IssueId, CommonConfiguration.TestCaseId }, linkConfiguration.LinkValueType, "Issue", false);

            // assert exception
        }

        /// <summary>
        /// A test for AddLinks
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_AddLinkWithOverwriteFalse_ShouldAddLink()
        {
            var linkConfiguration = GetRequirementConfigurationLinkItemForLinkType(LinkType.Affects);

            var old = _requirement.Links.First(x => x.Key.LinkValueType == "Affects").Value.Length;
            Assert.AreNotEqual(0, old);
            _requirement.AddLinks(_adapter, new int[] { }, linkConfiguration.LinkValueType, false);
            Assert.AreEqual(old, _requirement.Links.First(x => x.Key.LinkValueType == "Affects").Value.Length);
            Assert.IsFalse(_requirement.IsDirty);
        }

        /// <summary>
        /// A test for GetRevision
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_ChangePlainTextField_ShouldBeSaved()
        {
            _requirement.Fields["Microsoft.VSTS.Common.Triage"].Value = "Triaged";
            _adapter.Save();

            _requirement.Fields["Microsoft.VSTS.Common.Triage"].Value = "Pending";
            Assert.AreEqual("Pending", _requirement.Fields["Microsoft.VSTS.Common.Triage"].Value);
            _adapter.Save();

            var adapter = new Tfs2012SyncAdapter(_testServerConfiguration.TeamProjectCollectionUrl, _testServerConfiguration.TeamProjectName, null, _testServerConfiguration.Configuration);

            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.RequirementId }));
            var requirement = (TfsWorkItem)adapter.WorkItems.Find(CommonConfiguration.RequirementId);

            Assert.AreEqual("Pending", requirement.Fields["Microsoft.VSTS.Common.Triage"].Value);
        }

        /// <summary>
        /// Tests whether html code can be written and read and whether embedded images are automatically attached
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Keep test case simple.")]
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_SetValueHTML_ShouldAttachHTMLImages()
        {
            var filename = Path.GetTempPath() + "AIT\\image.jpg";
            var dirName = Path.GetDirectoryName(filename);
            if (dirName == null)
                Assert.Fail("File does not exist.");
            Directory.CreateDirectory(dirName);

            using (var streamWriter = new StreamWriter(File.Open(filename, FileMode.Create)))
            {
                streamWriter.Write("sdfdsf");
            }

            _requirement.WorkItem.Attachments.Clear();
            var html = string.Format(CultureInfo.InvariantCulture, "<html><body><span>FOO <img src=\"{0}\"/> BAR</span></body></html>", filename);
            _requirement.Fields[_testServerConfiguration.HtmlFieldReferenceName].Value = html;
            Assert.IsTrue(_requirement.IsDirty);
            Assert.AreEqual(0, _adapter.Save().Count);

            var adapter = new Tfs2012SyncAdapter(_testServerConfiguration.TeamProjectCollectionUrl, _testServerConfiguration.TeamProjectName, null, _testServerConfiguration.Configuration);

            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.RequirementId }));
            var requirement = (TfsWorkItem)adapter.WorkItems.Find(CommonConfiguration.RequirementId);
            Assert.IsTrue(requirement.Fields[_testServerConfiguration.HtmlFieldReferenceName].Value.Contains("FOO"));
            Assert.IsTrue(requirement.Fields[_testServerConfiguration.HtmlFieldReferenceName].Value.Contains("BAR"));

            Assert.AreEqual(1, requirement.WorkItem.AttachedFileCount);
        }

        /// <summary>
        /// A test for GetRevision
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Keep test case simple.")]
        public void TfsWorkItem_GetSetMicroDocument_ShouldAttachMicroDocumentAndLoadIt()
        {
            var filename = Path.GetTempPath() + "AIT\\System.Description.docx";
            var dirName = Path.GetDirectoryName(filename);
            if (dirName == null)
                Assert.Fail("File does not exist.");
            Directory.CreateDirectory(dirName);
            var content = Guid.NewGuid();
            using (var streamWriter = new StreamWriter(File.Open(filename, FileMode.Create)))
            {
                streamWriter.Write(content);
            }

            _issue.Fields["System.Description"].MicroDocument = filename;
            Assert.IsTrue(_issue.IsDirty);
            _adapter.Save();
            Assert.IsTrue(_issue.WorkItem.AttachedFileCount > 0);

            var adapter = new Tfs2012SyncAdapter(_testServerConfiguration.TeamProjectCollectionUrl, _testServerConfiguration.TeamProjectName, null, _testServerConfiguration.Configuration);
            Assert.IsTrue(adapter.Open(new[] { CommonConfiguration.IssueId }));

            var issue = (TfsWorkItem)adapter.WorkItems.Find(CommonConfiguration.IssueId);

            Assert.IsTrue(File.Exists((string)issue.Fields["System.Description"].MicroDocument));

            var actualContent = File.ReadAllLines((string)issue.Fields["System.Description"].MicroDocument);
            Assert.AreEqual(content.ToString(), actualContent[0]);
        }

        /// <summary>
        /// A test for IsDirty
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_MakeChanges_ValidateIsDirty()
        {
            Assert.IsFalse(_requirement.IsDirty);
            _requirement.Fields["System.Title"].Value = "Title" + Guid.NewGuid();
            Assert.IsTrue(_requirement.IsDirty);
            Assert.AreEqual(0, _adapter.Save().Count);
            Assert.IsFalse(_requirement.IsDirty);
            _requirement.AddLinks(_adapter, new int[] { }, "Child", true);
            _adapter.Save();
            _requirement.AddLinks(_adapter, new[] { CommonConfiguration.LinkTestWorkItemId }, "Child", true);
            Assert.IsTrue(_requirement.IsDirty);
            _adapter.Save();
            Assert.IsFalse(_requirement.IsDirty);
        }

        /// <summary>
        /// A test for IsNew
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsWorkItem_CreateWorkItems_ShouldBeNew()
        {
            Assert.IsTrue(_adapter.CreateNewWorkItem(_testServerConfiguration.Configuration.GetWorkItemConfiguration("Requirement")).IsNew);
        }

        /// <summary>
        /// A test for Links
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsWorkItem_Links_Validate()
        {
            Assert.AreEqual(1, _requirement.Links[_requirement.Links.Keys.First(x => x.LinkValueType == "Affects")].Length);
            Assert.AreEqual(CommonConfiguration.RequirementAffectedItemId, _requirement.Links[_requirement.Links.Keys.First(x => x.LinkValueType == "Affects")].ElementAt(0)); // Id of the linked item
        }

        #endregion Tests

        #region Helper members

        private enum LinkType
        {
            Child,
            Affects,
        }

        private int[] GetIdsOfLinksOfGivenLinkTypeFromWorkItem(IConfigurationLinkItem linkConfiguration, IWorkItem workItem)
        {
            return workItem.Links.First(x => x.Key.LinkValueType == linkConfiguration.LinkValueType).Value;
        }

        private IConfigurationLinkItem GetRequirementConfigurationLinkItemForLinkType(LinkType linkType)
        {
            return _testServerConfiguration.Configuration.GetWorkItemConfiguration("Requirement").Links.First(x => x.LinkValueType == linkType.ToString());
        }

        private void RemoveExistingLinksOfGivenLinkTypeFromWorkItem(IConfigurationLinkItem linkConfiguration, IWorkItem workItem)
        {
            workItem.AddLinks(
                _adapter,
                new int[] { },
                linkConfiguration.LinkValueType,
                true);
            InvalidateTestIfNotEqual(0, GetIdsOfLinksOfGivenLinkTypeFromWorkItem(linkConfiguration, workItem).Length, "workItem.AddLinks() did not deleted all links successfully.");
        }

        #endregion Helper members

        #region Test invalidations on errors in arrange parts

        /// <summary>
        /// Marks the test as inconclusive if the specified condition is false.
        /// Use this method to distinguish an error in an arrange part of a test from an error the in call under test.
        /// </summary>
        /// <param name="condition">The condition of an arrange part to verify is true.</param>
        /// <param name="message">A message to display. This message can be seen in the unit test results.</param>
        private static void InvalidateTestIfNotTrue(bool condition, string message)
        {
            if (condition)
                return;

            throw new AssertInconclusiveException($"There was an error in the test arrangement.{Environment.NewLine}IsTrue failed. {message}");
        }

        /// <summary>
        /// Marks the test as inconclusive if the two specified generic type data are not equal.
        /// Use this method to distinguish an error in an arrange part of a test from an error the in call under test.
        /// </summary>
        /// <param name="expected">The first generic type data to compare.</param>
        /// <param name="actual">The second generic type data to compare.</param>
        /// <param name="message">A message to display. This message can be seen in the unit test results.</param>
        private static void InvalidateTestIfNotEqual<T>(T expected, T actual, string message)
        {
            if ((expected == null && actual == null) || (expected != null && expected.Equals(actual)))
                return;

            throw new AssertInconclusiveException(
                $"There was an error in the test arrangement.{Environment.NewLine}AreEqual failed. Expected:<{(expected == null ? "null" : expected.ToString())}>. Actual:<{(actual == null ? "null" : actual.ToString())}>. {message}");
        }

        #endregion Test invalidations on errors in arrange parts
    }
}
