#region Usings
using System;
using System.Collections.Generic;
using System.Linq;
using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.TestReport;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using TFS.Test.Common;
#endregion
// ReSharper disable InconsistentNaming
// ReSharper disable RedundantCast
// ReSharper disable UnusedVariable
namespace TFS.SyncService.Model.Tests.Integration
{
    /// <summary>
    /// All tests to validate the correct function of the object queries
    /// </summary>

    [TestClass]
    [DeploymentItem("Microsoft.WITDataStore32.dll")]
    [DeploymentItem("Microsoft.WITDataStore64.dll")]
    public class TestReportHelperObjectQueryTest
    {

        #region Private properties

        private static ITestManagementTeamProject testManagement;
        private static TfsTeamProjectCollection projectCollection;
        private static ITestPlanCollection testPlans;
        private static IEnumerable<ITestCase> testCases;
        private static ITestSuiteCollection testSuites;
        private static ITfsTestAdapter _testAdapter;
        private static ITfsTestPlanDetail _testPlanDetail;
        private static ITfsTestSuiteDetail _testSuite_1_1Detail;
        private static ITfsTestCaseDetail _testCaseDetailAllSuitesAndPlans;
        private static ITfsTestCaseDetail _testCaseDetailForTestCaseInSuite1_2_1;
        private static IList<ITfsPropertyValueProvider> _resultForTestPlan;
        private static IList<ITfsPropertyValueProvider> _resultForTestSuite;
        private static IList<ITfsPropertyValueProvider> _resultForTestCaseInAllSuitesAndPlans;
        private static IList<ITfsPropertyValueProvider> _resultForTestCaseInSuite1_2_1;
        private static TestContext _testContext;
        private const string _checkName = "Name";
        private const string _checkId = "Id";

        #endregion

        #region Test Initializations

        /// <summary>
        /// Make sure all services are registered.
        /// </summary>
        [ClassInitialize]
        public static void MyClassInitialize(TestContext testContext)
        {
            _testContext = testContext;
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);

            //Init the assemblies
            AIT.TFS.SyncService.Service.AssemblyInit.Instance.Init();
            AIT.TFS.SyncService.Adapter.TFS2012.AssemblyInit.Instance.Init();
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfiguration.TeamProjectCollectionUrl));
            testManagement = projectCollection.GetService<ITestManagementService>().GetTeamProject(serverConfiguration.TeamProjectName);

            InitializeTestData();
        }

        /// <summary>
        /// Initialize Test Adapter needed for all Test Methods
        /// </summary>
        [TestInitialize]
        public void MyTestInitialize()
        {
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");
            var serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            _testAdapter = SyncServiceFactory.CreateTfsTestAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, config);
        }

        /// <summary>
        ///Initiaze and get the test objects from the test managment.
        /// </summary>
        private static void InitializeTestData()
        {
            testCases = testManagement.TestCases.Query("SELECT * FROM WorkItems");
            testSuites = testManagement.TestSuites.Query("SELECT * FROM TestSuite");
            testPlans = testManagement.TestPlans.Query("SELECT * FROM TestPlan");
        }

        #endregion Test Initializations

        #region Test Methods

        /// <summary>
        /// Test
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        [DeploymentItem("SharedTestFiles", "SharedTestFiles")]
        public void TestReportHelper_GetLinkedConfigurationForTestManagementObject_ShouldReturnCorrectConfigurations()
        {
            // Arrange
            const bool includeOnlyMostRecentTestResults = false;
            const bool includeOnlyMostRecentTestResultForAllConfigurations = false;
            InitializeTestPlanStructure();

            // Act
            ExecuteObjectQuery(includeOnlyMostRecentTestResults, includeOnlyMostRecentTestResultForAllConfigurations);

            //Assert
            // Check results for TestPlan, should contain configurations ids 2,3,4,5,6,7
            Assert.IsNotNull(_resultForTestPlan);
            Assert.AreEqual(7, _resultForTestPlan.Count);

            var item = _resultForTestPlan;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_1_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_2_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_2_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "2"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "3"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "4"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "5"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "6"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "7"));

            //Check results for TestSuite, should contain configurations ids 2,3,5
            Assert.IsNotNull(_resultForTestSuite);
            Assert.AreEqual(7, _resultForTestSuite.Count);

            item = _resultForTestSuite;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_1_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "2"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "3"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "5"));

            //Check results for Testcase that is in all configurations,  should contain configurations ids 2
            Assert.IsNotNull(_resultForTestCaseInAllSuitesAndPlans);
            Assert.AreEqual(1, _resultForTestCaseInAllSuitesAndPlans.Count);

            item = _resultForTestCaseInAllSuitesAndPlans;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "2"));

            //Check results for Testcase that is in the test suite 1_2_1. This TC should contain configurations 2 and 6
            Assert.IsNotNull(_resultForTestCaseInSuite1_2_1);
            Assert.AreEqual(2, _resultForTestCaseInSuite1_2_1.Count);

            item = _resultForTestCaseInSuite1_2_1;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_2_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "2"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "6"));
        }

        /// <summary>
        /// Test
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        [DeploymentItem("SharedTestFiles", "SharedTestFiles"), TestCategory("PreparationNeeded")]
        public void TestReportHelper_GetLinkedConfigurationForTestManagementObject_OnlyForLatestTestResults_ShouldReturnCorrectConfigurations()
        {
            // Arrange
            const bool includeOnlyMostRecentTestResults = true;
            const bool includeOnlyMostRecentTestResultForAllConfigurations = false;
            InitializeTestPlanStructure();

            // Act
            ExecuteObjectQuery(includeOnlyMostRecentTestResults, includeOnlyMostRecentTestResultForAllConfigurations);

            // Assert
            // Check results for TestPlan, should contain configurations ids 6
            Assert.IsNotNull(_resultForTestPlan);
            Assert.AreEqual(1, _resultForTestPlan.Count);

            var item = _resultForTestPlan;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_2_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "6"));

            // Check results for TestSuite 1_1, here the latest test result should have the configuration 1_1 assigned
            Assert.IsNotNull(_resultForTestSuite);
            Assert.AreEqual(1, _resultForTestSuite.Count);

            item = _resultForTestSuite;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));
            Assert.IsFalse(HasTestItemPropertyValue(item, _checkName, "1_1 Configuration"));
            Assert.IsFalse(HasTestItemPropertyValue(item, _checkName, "1_1_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "2"));
            Assert.IsFalse(HasTestItemPropertyValue(item, _checkId, "3"));
            Assert.IsFalse(HasTestItemPropertyValue(item, _checkId, "5"));

            //Check results for Testcase that is in all configurations, should not contain any configurations
            Assert.IsNotNull(_resultForTestCaseInAllSuitesAndPlans);
            Assert.AreEqual(1, _resultForTestCaseInAllSuitesAndPlans.Count);

            item = _resultForTestCaseInAllSuitesAndPlans;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "2"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));

            //Check results for Testcase that is in the test suite 1_2_1. One result --> one configuration
            Assert.IsNotNull(_resultForTestCaseInSuite1_2_1);
            Assert.AreEqual(1, _resultForTestCaseInSuite1_2_1.Count);

            item = _resultForTestCaseInSuite1_2_1;
            Assert.IsFalse(HasTestItemPropertyValue(item, _checkId, "2"));
            Assert.IsFalse(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "6"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_2_1 Configuration"));
        }

        /// <summary>
        /// Test
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TestReportHelper_GetLinkedConfigurationForTestManagementObject_OnlyForLatestTestResultsAndAllConfigurations_ShouldReturnCorrectConfigurations()
        {
            // Arrange
            const bool includeOnlyMostRecentTestResults = true;
            const bool includeOnlyMostRecentTestResultForAllConfigurations = true;
            InitializeTestPlanStructure();

            // Act
            ExecuteObjectQuery(includeOnlyMostRecentTestResults, includeOnlyMostRecentTestResultForAllConfigurations);

            // Assert
            // Check results for TestPlan, should contain configurations ids 2,5,6
            Assert.IsNotNull(_resultForTestPlan);
            Assert.AreEqual(7, _resultForTestPlan.Count);

            var item = _resultForTestPlan;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_1_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_2_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "2"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "5"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "6"));

            //Check results for TestSuite, should contain configurations ids 2,5
            Assert.IsNotNull(_resultForTestSuite);
            Assert.AreEqual(7, _resultForTestSuite.Count);

            item = _resultForTestSuite;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_1_1 Configuration"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "2"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "5"));

            //Check results for Testcase that is in all configurations,  should contain configurations ids 2
            Assert.IsNotNull(_resultForTestCaseInAllSuitesAndPlans);
            Assert.AreEqual(1, _resultForTestCaseInAllSuitesAndPlans.Count);

            item = _resultForTestCaseInAllSuitesAndPlans;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "2"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "ConfigurationThatShouldAppearInAllTestCases"));

            //Check results for Testcase that is in the test suite 1_2_1. This TC should only 6
            Assert.IsNotNull(_resultForTestCaseInSuite1_2_1);
            Assert.AreEqual(1, _resultForTestCaseInSuite1_2_1.Count);

            item = _resultForTestCaseInSuite1_2_1;
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkId, "6"));
            Assert.IsTrue(HasTestItemPropertyValue(item, _checkName, "1_2_1 Configuration"));
        }
        #endregion Test Methods

        #region Test Helper Methods

        private static void InitializeTestPlanStructure()
        {
            var testPlan = testPlans.FirstOrDefault(x => x.Id == CommonConfiguration.TestPlan_1_Id);
            var testSuite_1_1 = testSuites.FirstOrDefault(x => x.Id == CommonConfiguration.TestSuite_1_1_TestSuiteId);
            var testSuite_1_2_1 = testSuites.FirstOrDefault(x => x.Id == CommonConfiguration.TestSuite_1_2_1TestSuiteId);
            var testCaseInAllTestSuitesAndPlans = testCases.FirstOrDefault(x => x.Id == CommonConfiguration.TestCase_InAllSuitesAndPlans_Id);
            var testCaseOnlyInSuite1_2_1 = testCases.FirstOrDefault(x => x.Id == CommonConfiguration.TestCase_1_2_1_Id);

            var testPlanWrapper = new TfsTestPlan((ITestPlan)testPlan);
            _testPlanDetail = new TfsTestPlanDetail(testPlanWrapper);

            var testSuite_1_1Wrapper = new TfsTestSuite((IStaticTestSuite)testSuite_1_1, testPlanWrapper);
            _testSuite_1_1Detail = new TfsTestSuiteDetail(testSuite_1_1Wrapper);

            var testSuite_1_2_1Wrapper = new TfsTestSuite((IStaticTestSuite)testSuite_1_2_1, testPlanWrapper);

            var testCaseWrapperAllSuitesAndPlans = new TfsTestCase((ITestCase)testCaseInAllTestSuitesAndPlans, testSuite_1_1Wrapper);
            _testCaseDetailAllSuitesAndPlans = new TfsTestCaseDetail(testCaseWrapperAllSuitesAndPlans);

            var testCaseWrapperForTestCaseInSuite1_2_1 = new TfsTestCase((ITestCase)testCaseOnlyInSuite1_2_1, testSuite_1_2_1Wrapper);
            _testCaseDetailForTestCaseInSuite1_2_1 = new TfsTestCaseDetail(testCaseWrapperForTestCaseInSuite1_2_1);
        }

        private static void ExecuteObjectQuery(bool includeOnlyMostRecentTestResults, bool includeOnlyMostRecentTestResultForAllConfigurations)
        {
            var destElement = new Mock<IElement>();
            var objectQuery = new Mock<IObjectQuery>();
            objectQuery.SetupGet(x => x.DestinationElements).Returns(new List<IElement> { destElement.Object });
            objectQuery.SetupGet(x => x.Name).Returns("TestObjectQuery");

            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupAllProperties();

            var replacement = new Mock<IConfigurationTestReplacement>();
            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(x => x.ConfigurationTest.SetHyperlinkBase).Returns(true);
            configuration.SetupGet(x => x.ConfigurationTest.ShowHyperlinkBaseMessageBoxes).Returns(false);

            destElement.SetupGet(x => x.ItemName).Returns("Configurations");

            var helper = new TestReportHelper(_testAdapter, wordTestAdapter.Object, configuration.Object, () => false);
            if (includeOnlyMostRecentTestResults)
                helper.IncludeMostRecentTestResults = true;
            if (includeOnlyMostRecentTestResultForAllConfigurations)
                helper.IncludeOnlyMostRecentTestResultForAllConfigurations = true;

            _resultForTestCaseInAllSuitesAndPlans = helper.GetLinkedConfigurationForTestManagementObject(_testCaseDetailAllSuitesAndPlans, objectQuery.Object, false);
            _resultForTestCaseInSuite1_2_1 = helper.GetLinkedConfigurationForTestManagementObject(_testCaseDetailForTestCaseInSuite1_2_1, objectQuery.Object, false);
            _resultForTestPlan = helper.GetLinkedConfigurationForTestManagementObject(_testPlanDetail, objectQuery.Object, false);
            _resultForTestSuite = helper.GetLinkedConfigurationForTestManagementObject(_testSuite_1_1Detail, objectQuery.Object, false);
        }

        private static bool HasTestItemPropertyValue(IList<ITfsPropertyValueProvider> item, string propertyName, string propertyValue)
        {
            return item.Any(x => x.PropertyValue(propertyName, _testAdapter.ExpandEnumerable).Equals(propertyValue));
        }
        #endregion Test Helper Methods
    }
}
