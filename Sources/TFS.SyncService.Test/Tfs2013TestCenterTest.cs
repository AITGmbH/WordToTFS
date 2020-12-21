namespace TFS.SyncService.Test.Unit
{
    #region Usings
    using System.Linq;
    using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
    using AIT.TFS.SyncService.Contracts.Enums;
    using AIT.TFS.SyncService.Contracts.Exceptions;
    using AIT.TFS.SyncService.Contracts.TestCenter;
    using AIT.TFS.SyncService.Factory;
    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.TestManagement.Client;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TFS.Test.Common;
    #endregion

    /// <summary>
    /// Test cases for the test center
    /// </summary>
    [TestClass]
    public class Tfs20123TestCenterTest
    {
        #region Fields
        private static ITestManagementTeamProject testManagement;
        private static ITfsTestAdapter testAdapter;
        #endregion

        #region Test initializations

        /// <summary>
        /// Make sure all service are registered.
        /// </summary>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            AIT.TFS.SyncService.Service.AssemblyInit.Instance.Init();
            AIT.TFS.SyncService.Adapter.TFS2012.AssemblyInit.Instance.Init();
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(CommonConfiguration.Tfs2013ServerConfiguration(testContext).ServerName, true, false);
            testManagement = projectCollection.GetService<ITestManagementService>().GetTeamProject(CommonConfiguration.Tfs2013ServerConfiguration(testContext).ProjectName);
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");
            testAdapter = SyncServiceFactory.CreateTfsTestAdapter(CommonConfiguration.Tfs2013ServerConfiguration(testContext).ServerName, CommonConfiguration.Tfs2013ServerConfiguration(testContext).ProjectName, config);
        }
        #endregion

        #region TestMethods

        /// <summary>
        /// Testing if the "WorkItem" fields of a test case are accessible through the test management. The result is compared with the original work item.
        /// </summary>
        [TestMethod]
        public void TestCenter_TestCase_WorkItemFieldsShouldBeAccessible()
        {
            // arrange
            var testCases = testManagement.TestCases.Query("SELECT * FROM WorkItems");
            var microsoftTestCase = testCases.FirstOrDefault(testCase => testCase.Id == CommonConfiguration.TestCase_1_1_1_Id);
            var tfsTestCase = new TfsTestCase(microsoftTestCase);

            // act
            var tfsTestCaseDetail = new TfsTestCaseDetail(tfsTestCase);

            // assert
            // ReSharper disable once PossibleNullReferenceException
            Assert.AreEqual(
                microsoftTestCase.Id.ToString(),
                tfsTestCaseDetail.PropertyValue("WorkItem.Id", testAdapter.ExpandEnumerable),
                "WorkItem.Id incorrect.");
            Assert.AreEqual(
                microsoftTestCase.Title,
                tfsTestCaseDetail.PropertyValue("WorkItem.Title", testAdapter.ExpandEnumerable),
                "WorkItem.Title incorrect.");
            Assert.AreEqual(
                microsoftTestCase.Revision.ToString(),
                tfsTestCaseDetail.PropertyValue("WorkItem.Rev", testAdapter.ExpandEnumerable),
                "WorkItem.Rev incorrect.");
        }

        /// <summary>
        /// Testing if custom fields of a test case are accessible through the test management. The result is compared with the original work item.
        /// </summary>
        [TestMethod]
        public void TestCenter_TestCase_CustomFieldsShouldBeAccessible()
        {
            // arrange
            var testCases = testManagement.TestCases.Query("SELECT * FROM WorkItems");
            var microsoftTestCase = testCases.FirstOrDefault(testCase => testCase.Id == CommonConfiguration.TestCase_1_1_1_Id);
            var tfsTestCase = new TfsTestCase(microsoftTestCase);
            
            // act
            var tfsTestCaseDetail = new TfsTestCaseDetail(tfsTestCase);
            
            // assert
            // ReSharper disable once PossibleNullReferenceException
            Assert.AreEqual(
                microsoftTestCase.CustomFields["Microsoft.VSTS.TCM.AutomationStatus"].Value.ToString(),
                tfsTestCaseDetail.PropertyValue("CustomFields[\"Microsoft.VSTS.TCM.AutomationStatus\"].Value", testAdapter.ExpandEnumerable),
                "CustomFields[\"Microsoft.VSTS.TCM.AutomationStatus\"].Value incorrect.");
        }

        /// <summary>
        /// Testing if the "WorkItem" fields of a test plan are accessible through the test management. The result is compared with the original work item.
        /// </summary>
        [TestMethod]
        public void TestCenter_TestPlan_WorkItemFieldsShouldBeAccessible()
        {
            // arrange
            var testPlans = testManagement.TestPlans.Query("SELECT * FROM TestPlan");
            var microsoftTestPlan = testPlans.FirstOrDefault(testPlan => testPlan.Id == CommonConfiguration.TestPlan_1_Id);
            var tfsTestPlan = new TfsTestPlan(microsoftTestPlan);

            // act
            var tfsTestPlanDetail = new TfsTestPlanDetail(tfsTestPlan);

            // assert
            // ReSharper disable once PossibleNullReferenceException
            try
            {
                Assert.AreEqual(
                    microsoftTestPlan.Id.ToString(),
                    tfsTestPlanDetail.PropertyValue("WorkItem.Id", testAdapter.ExpandEnumerable),
                    "WorkItem.Id incorrect.");
                Assert.AreEqual(
                    microsoftTestPlan.Name,
                    tfsTestPlanDetail.PropertyValue("WorkItem.Name", testAdapter.ExpandEnumerable),
                    "WorkItem.Name incorrect.");
            }
            catch (ConfigurationException)
            {
                Assert.Inconclusive("Test plans are not yet treated as work items.");
            }
            Assert.Inconclusive("Test plans are treated as work items. Remove Inconclusives & try-catch from this test.");

        }
        #endregion
    }
}