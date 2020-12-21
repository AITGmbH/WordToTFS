#region Usings
using System;
using System.Collections.Generic;
using System.Linq;
using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Factory;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TFS.Test.Common;
#endregion

// ReSharper disable InconsistentNaming

namespace TFS.SyncService.Test.Unit
{
    /// <summary>
    /// Test cases for the test center
    /// </summary>
    [TestClass]
    public class TfsTestCenterTest
    {
        private static ITestConfiguration _testConfiguration;
        private static ITestManagementTeamProject _testManagement;
        private static ITfsTestAdapter _testAdapter;
        private static ISharedStep _sharedStep;
        private static ITestCase _testCase;
        private static TestContext _testContext;

        //#region Test initializations

        /// <summary>
        /// Make sure all service are registered.
        /// </summary>
        [ClassInitialize]
        [TestCategory("ConnectionNeeded")]
        public static void MyClassInitialize(TestContext testContext)
        {
            _testContext = testContext;

            AIT.TFS.SyncService.Service.AssemblyInit.Instance.Init();
            AIT.TFS.SyncService.Adapter.TFS2012.AssemblyInit.Instance.Init();

            var serverConfig = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);
            var config = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title");

            _testAdapter = SyncServiceFactory.CreateTfsTestAdapter(serverConfig.TeamProjectCollectionUrl, serverConfig.TeamProjectName, config);
            var  projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverConfig.TeamProjectCollectionUrl));
            _testManagement = projectCollection.GetService<ITestManagementService>().GetTeamProject(serverConfig.TeamProjectName);

            var sharedSteps = _testManagement.SharedSteps.Query("SELECT * FROM WorkItems WHERE System.Title='UnitTest_SharedStep'");
            _sharedStep = sharedSteps.FirstOrDefault();

            if (_sharedStep == null)
            {
                _sharedStep = _testManagement.SharedSteps.Create();
                var sharedStep1 = _sharedStep.CreateTestStep();
                sharedStep1.Title = "First Shared Step";
                sharedStep1.ExpectedResult = "Result of first shared step";

                var sharedStep2 = _sharedStep.CreateTestStep();
                sharedStep2.Title = "Second Shared Step .: @ParametersDontLikeSpecialChars";
                sharedStep2.ExpectedResult = "Result of second shared step";

                _sharedStep.Actions.Add(sharedStep1);
                _sharedStep.Actions.Add(sharedStep2);

                _sharedStep.Title = "UnitTest_SharedStep_1";
                _sharedStep.Save();
            }
            else
            {
                var sharedStep1 = _sharedStep.Actions[0] as ITestStep;
                if (sharedStep1 != null)
                {
                    sharedStep1.Title = "First Shared Step";
                    sharedStep1.ExpectedResult = "Result of first shared step";
                }

                var sharedStep2 = _sharedStep.Actions[1] as ITestStep;
                if (sharedStep2 != null)
                {
                    sharedStep2.Title = "Second Shared Step .: @ParametersDontLikeSpecialChars";
                    sharedStep2.ExpectedResult = "Result of second shared step";
                }

                _sharedStep.WorkItem.Open();
                _sharedStep.Save();
            }

            var testCases = _testManagement.TestCases.Query("SELECT * FROM WorkItems WHERE System.Title='UnitTest_TestCase'");
            _testCase = testCases.FirstOrDefault();

            if (_testCase == null)
            {
                _testCase = _testManagement.TestCases.Create();
                _testCase.Title = "UnitTest_TestCase";

                var step1 = _testCase.CreateTestStep();
                step1.Title = "First Step";
                step1.ExpectedResult = "Result of first step";

                var step2 = _testCase.CreateTestStep();
                step2.Title = "Second Step";
                step2.ExpectedResult = "Result of second step";

                var step3 = _testCase.CreateTestStep();
                step3.Title = "@DescriptionParameter";
                step3.ExpectedResult = "@ExpectedResultParameter";

                var ssr = _testCase.CreateSharedStepReference();
                ssr.SharedStepId = _sharedStep.Id;

                _testCase.Actions.Add(step1);
                _testCase.Actions.Add(step2);
                _testCase.Actions.Add(ssr);
                _testCase.Actions.Add(step3);
                _testCase.Save();
            }

            var testConfigurations = _testManagement.TestConfigurations.Query("Select * from TestConfiguration where Name='UnitTest_TestConfiguration'");
            _testConfiguration = testConfigurations.FirstOrDefault();

            if (_testConfiguration == null)
            {
                _testConfiguration = _testManagement.TestConfigurations.Create();
                _testConfiguration.Name = "UnitTest_TestConfiguration";
                _testConfiguration.Save();
            }
        }

        /// <summary>
        /// Testing if the "WorkItem" fields of a test case are accessible through the test management. The result is compared with the original work item.
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TestCenter_TestCase_WorkItemFieldsShouldBeAccessible()
        {
            // arrange
            var testCases = _testManagement.TestCases.Query("SELECT * FROM WorkItems");
            var microsoftTestCase = testCases.FirstOrDefault(testCase => testCase.Id == CommonConfiguration.TestCase_1_1_1_Id);
            var tfsTestCase = new TfsTestCase(microsoftTestCase);

            // act
            var tfsTestCaseDetail = new TfsTestCaseDetail(tfsTestCase);

            // assert
            // ReSharper disable once PossibleNullReferenceException
            Assert.AreEqual(
                microsoftTestCase.Id.ToString(),
                tfsTestCaseDetail.PropertyValue("WorkItem.Id", _testAdapter.ExpandEnumerable),
                "WorkItem.Id incorrect.");
            Assert.AreEqual(
                microsoftTestCase.Title,
                tfsTestCaseDetail.PropertyValue("WorkItem.Title", _testAdapter.ExpandEnumerable),
                "WorkItem.Title incorrect.");
            Assert.AreEqual(
                microsoftTestCase.Revision.ToString(),
                tfsTestCaseDetail.PropertyValue("WorkItem.Rev", _testAdapter.ExpandEnumerable),
                "WorkItem.Rev incorrect.");
        }

        /// <summary>
        /// Testing if custom fields of a test case are accessible through the test management. The result is compared with the original work item.
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TestCenter_TestCase_CustomFieldsShouldBeAccessible()
        {
            // arrange
            var testCases = _testManagement.TestCases.Query("SELECT * FROM WorkItems");
            var microsoftTestCase = testCases.FirstOrDefault(testCase => testCase.Id == CommonConfiguration.TestCase_1_1_1_Id);
            var tfsTestCase = new TfsTestCase(microsoftTestCase);

            // act
            var tfsTestCaseDetail = new TfsTestCaseDetail(tfsTestCase);

            // assert
            // ReSharper disable once PossibleNullReferenceException
            Assert.AreEqual(
                microsoftTestCase.CustomFields["Microsoft.VSTS.TCM.AutomationStatus"].Value.ToString(),
                tfsTestCaseDetail.PropertyValue("CustomFields[\"Microsoft.VSTS.TCM.AutomationStatus\"].Value", _testAdapter.ExpandEnumerable),
                "CustomFields[\"Microsoft.VSTS.TCM.AutomationStatus\"].Value incorrect.");
        }

       // #region TestMethods

        /// <summary>
        /// Test action should provide a property "StepNumber" that returns the position of the step as defined in the test case
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestAction_ShouldProvideStepNumberProperty()
        {
            var testCaseWrapper = new TfsTestCase(_testCase);
            var testCaseDetail = new TfsTestCaseDetail(testCaseWrapper);

            // numberings starts at "1"
            Assert.AreEqual("1", testCaseDetail.PropertyValue("Actions[0].StepNumber", x => x));
        }

        /// <summary>
        /// In a test case references shared test steps, they should be inline and have an step number like "3.1"
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestAction_ShouldCreateStepNumberHierarchyForSharedTestSteps()
        {
            var testCaseWrapper = new TfsTestCase(_testCase);
            var testCaseDetail = new TfsTestCaseDetail(testCaseWrapper, true);

            Assert.AreEqual("3.1", testCaseDetail.PropertyValue("Actions[3].StepNumber", x => x));
        }

        /// <summary>
        /// In a test case references shared test steps, the first step before the shared step, should only contan the number
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestAction_ShouldCreateHeadingForSharedSteps()
        {
            var testCaseWrapper = new TfsTestCase(_testCase);
            var testCaseDetail = new TfsTestCaseDetail(testCaseWrapper, true);

            Assert.AreEqual("3.0", testCaseDetail.PropertyValue("Actions[2].StepNumber", x => x));
        }

        /// <summary>
        /// TestResults should provide property for step number
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestActionResult_ShouldProvideTestNumberProperty()
        {
            var testRun = _testManagement.TestRuns.Create();

            testRun.AddTest(_testCase.Id, _testConfiguration.Id, _testManagement.TfsIdentityStore.FindByAccountName("Administrator"));
            testRun.Save();

            var testResult = testRun.QueryResults()[0];
            var iteration = testResult.CreateIteration(1);
            testResult.Iterations.Add(iteration);

            var stepResult = iteration.CreateStepResult(_testCase.Actions[0].Id);
            iteration.Actions.Add(stepResult);

            testRun.Save();

            var testCaseWrapper = new TfsTestResult(testResult, _testCase.Id);
            var testCaseDetail = new TfsTestResultDetail(testCaseWrapper, testCaseWrapper.OriginalTestResult.TestRunId);

            Assert.AreEqual("1", testCaseDetail.PropertyValue("Iterations[0].Actions[0].StepNumber", x => x));
        }

        /// <summary>
        /// If test case results references shared test steps, they should be inline and have an step number like "3.1"
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestActionResult_TestCenter_TestAction_ShouldCreateStepNumberHierarchyForSharedTestSteps()
        {
            var testRun = _testManagement.TestRuns.Create();

            testRun.AddTest(_testCase.Id, _testConfiguration.Id, _testManagement.TfsIdentityStore.FindByAccountName("Administrator"));
            testRun.Save();

            var testResult = testRun.QueryResults()[0];
            var iteration = testResult.CreateIteration(1);
            testResult.Iterations.Add(iteration);

            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[0].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[1].Id));

            var sharedStepResult = iteration.CreateSharedStepResult(_testCase.Actions[2].Id, _sharedStep.Id);
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[0].Id));
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[1].Id));
            iteration.Actions.Add(sharedStepResult);

            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[3].Id));

            testRun.Save();

            var testCaseWrapper = new TfsTestResult(testResult, _testCase.Id);
            var testCaseDetail = new TfsTestResultDetail(testCaseWrapper, testCaseWrapper.OriginalTestResult.TestRunId);

            Assert.AreEqual("3.2", testCaseDetail.PropertyValue("Iterations[0].Actions[3].StepNumber", x => x));
        }

        /// <summary>
        /// StepNumbers should be ordered as in the original test case
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestActionResult_ShouldAssignStepNumbersAsDefinedInTestCase()
        {
            var testRun = _testManagement.TestRuns.Create();

            testRun.AddTest(_testCase.Id, _testConfiguration.Id, _testManagement.TfsIdentityStore.FindByAccountName("Administrator"));
            testRun.Save();

            var testResult = testRun.QueryResults()[0];
            var iteration = testResult.CreateIteration(1);
            testResult.Iterations.Add(iteration);

            // note that these are shuffled
            var sharedStepResult = iteration.CreateSharedStepResult(_testCase.Actions[2].Id, _sharedStep.Id);
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[0].Id));
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[1].Id));
            iteration.Actions.Add(sharedStepResult);

            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[0].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[3].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[1].Id));

            testRun.Save();

            var testCaseWrapper = new TfsTestResult(testResult, _testCase.Id);
            var testCaseDetail = new TfsTestResultDetail(testCaseWrapper, testCaseWrapper.OriginalTestResult.TestRunId);

            Assert.AreEqual("3.2", testCaseDetail.PropertyValue("Iterations[0].Actions[3].StepNumber", x => x));
        }

        /// <summary>
        /// Should insert iteration variables into parameter
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestActionResult_ShouldInsertIterationVariablesIntoParameter()
        {
            var testRun = _testManagement.TestRuns.Create();

            testRun.AddTest(_testCase.Id, _testConfiguration.Id, _testManagement.TfsIdentityStore.FindByAccountName("Administrator"));
            testRun.Save();

            var testResult = testRun.QueryResults()[0];
            var iteration = testResult.CreateIteration(1);
            testResult.Iterations.Add(iteration);

            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[0].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[1].Id));

            var sharedStepResult = iteration.CreateSharedStepResult(_testCase.Actions[2].Id, _sharedStep.Id);
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[0].Id));
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[1].Id));
            iteration.Actions.Add(sharedStepResult);

            var testStepResult = iteration.CreateStepResult(_testCase.Actions[3].Id);
            testStepResult.Parameters.Add("DescriptionParameter", "MyExpectedDescriptionValue", "MyActualValue");
            testStepResult.Parameters.Add("ExpectedResultParameter", "MyExpectedResultValue", "MyActualValue");
            iteration.Actions.Add(testStepResult);

            testRun.Save();

            var testCaseWrapper = new TfsTestResult(testResult, _testCase.Id);
            var testCaseDetail = new TfsTestResultDetail(testCaseWrapper, testCaseWrapper.OriginalTestResult.TestRunId);

            Assert.AreEqual("MyExpectedDescriptionValue", testCaseDetail.PropertyValue("Iterations[0].Actions[4].Title", x => x));
            Assert.AreEqual("MyExpectedResultValue", testCaseDetail.PropertyValue("Iterations[0].Actions[4].ExpectedResult", x => x));
        }

        /// <summary>
        /// Should insert iteration variables into parameters. this fails if the parameterized string
        /// contains special characters like ":".
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestActionResult_ShouldInsertIterationVariablesIntoParameterWithSpecialChars()
        {
            var testRun = _testManagement.TestRuns.Create();

            testRun.AddTest(_testCase.Id, _testConfiguration.Id, _testManagement.TfsIdentityStore.FindByAccountName("Administrator"));
            testRun.Save();

            var testResult = testRun.QueryResults()[0];
            var iteration = testResult.CreateIteration(1);
            testResult.Iterations.Add(iteration);

            // note that these are shuffled
            var sharedStepResult = iteration.CreateSharedStepResult(_testCase.Actions[2].Id, _sharedStep.Id);
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[0].Id));
            var sharedStepActionResult2 = iteration.CreateStepResult(_sharedStep.Actions[1].Id);
            sharedStepActionResult2.Parameters.Add("ParametersDontLikeSpecialChars", "ShouldWorkDespiteSpecialChar", "MyActualValue");
            sharedStepResult.Actions.Add(sharedStepActionResult2);

            iteration.Actions.Add(sharedStepResult);
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[0].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[1].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[3].Id));

            testRun.Save();

            var testCaseWrapper = new TfsTestResult(testResult, _testCase.Id);
            var testCaseDetail = new TfsTestResultDetail(testCaseWrapper, testCaseWrapper.OriginalTestResult.TestRunId);

            Assert.AreEqual("Second Shared Step .: ShouldWorkDespiteSpecialChar", testCaseDetail.PropertyValue("Iterations[0].Actions[3].Title", x => x));
        }

        /// <summary>
        /// The custom property LinkedWorkItems should not be null
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestResult_ShouldContainPropertyLinkedWorkItemsForTestResult()
        {
            var testRun = _testManagement.TestRuns.Create();

            testRun.AddTest(_testCase.Id, _testConfiguration.Id, _testManagement.TfsIdentityStore.FindByAccountName("Administrator"));
            testRun.Save();

            var testResult = testRun.QueryResults()[0];
            var iteration = testResult.CreateIteration(1);
            testResult.Iterations.Add(iteration);

            // note that these are shuffled
            var sharedStepResult = iteration.CreateSharedStepResult(_testCase.Actions[2].Id, _sharedStep.Id);
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[0].Id));
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[1].Id));
            iteration.Actions.Add(sharedStepResult);

            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[0].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[3].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[1].Id));

            testRun.Save();

            var testCaseWrapper = new TfsTestResult(testResult, _testCase.Id);
            var testCaseDetail = new TfsTestResultDetail(testCaseWrapper, testCaseWrapper.OriginalTestResult.TestRunId);

            Assert.IsNotNull(testCaseDetail.PropertyValue("LinkedWorkItemsForTestResult",x=>x));
        }

        /// <summary>
        /// The custom property LinkedWorkItems should return a list of workitems
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_TestResult_PropertyLinkedWorkItemsForTestResultShouldBeList()
        {
            var testRun = _testManagement.TestRuns.Create();

            testRun.AddTest(_testCase.Id, _testConfiguration.Id, _testManagement.TfsIdentityStore.FindByAccountName("Administrator"));
            testRun.Save();

            var testResult = testRun.QueryResults()[0];
            var iteration = testResult.CreateIteration(1);
            testResult.Iterations.Add(iteration);

            // note that these are shuffled
            var sharedStepResult = iteration.CreateSharedStepResult(_testCase.Actions[2].Id, _sharedStep.Id);
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[0].Id));
            sharedStepResult.Actions.Add(iteration.CreateStepResult(_sharedStep.Actions[1].Id));
            iteration.Actions.Add(sharedStepResult);

            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[0].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[3].Id));
            iteration.Actions.Add(iteration.CreateStepResult(_testCase.Actions[1].Id));

            testRun.Save();

            var testCaseWrapper = new TfsTestResult(testResult, _testCase.Id);
            var testCaseDetail = new TfsTestResultDetail(testCaseWrapper, testCaseWrapper.OriginalTestResult.TestRunId);

            Assert.AreEqual(typeof(List<WorkItem>), testCaseDetail.LinkedWorkItemsForTestResult.GetType());
        }

        /// <summary>
        /// Test the assigment of a custom object
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TestCenter_CustomObjects_ShouldContainCustomPropertyWorkItemsForTestResult()
        {
            var testRun = _testManagement.TestRuns.Create();

            testRun.AddTest(_testCase.Id, _testConfiguration.Id, _testManagement.TfsIdentityStore.FindByAccountName("Administrator"));
            testRun.Save();

            var testResult = testRun.QueryResults()[0];
            var iteration = testResult.CreateIteration(1);
            testResult.Iterations.Add(iteration);

            testRun.Save();

            var testCaseWrapper = new TfsTestResult(testResult, _testCase.Id);

            var testCaseDetail = new TfsTestResultDetail(testCaseWrapper, testCaseWrapper.OriginalTestResult.TestRunId);

            var testInt = 2;
            testCaseDetail.AddCustomObjects("ASampleProperty",testInt);

            Assert.AreEqual("2", testCaseDetail.PropertyValue("ASampleProperty", x => x));
            Assert.AreEqual("2", testCaseDetail.PropertyValue("ASampleProperty.ToString()", x => x));
        }
    }
}