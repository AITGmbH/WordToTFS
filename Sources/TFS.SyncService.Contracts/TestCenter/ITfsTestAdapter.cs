using System;
using System.Collections;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.BuildCenter;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines the functionality of Microsoft test manager.
    /// Interface provides methods and properties to examine test environment in used project.
    /// </summary>
    /// <remarks>
    /// <code>
    /// http://blogs.msdn.com/b/duat_le/archive/2010/02/25/wiql-for-test.aspx
    /// http://blogs.microsoft.co.il/blogs/shair/archive/2010/07/06/tfs-api-part-27-test-plans-test-suites-test-cases-mapping.aspx
    /// http://msdn.microsoft.com/en-us/library/bb130306.aspx
    /// </code>
    /// </remarks>
    public interface ITfsTestAdapter
    {
        /// <summary>
        /// Gets the available test plans.
        /// </summary>
        IList<ITfsTestPlan> AvailableTestPlans { get; }

        /// <summary>
        /// Gets the available test configurations.
        /// </summary>
        IList<ITfsTestConfiguration> AvailableTestConfigurations { get; }

        /// <summary>
        /// Gets the available test management team project.
        /// </summary>
        ITestManagementTeamProject TestManagement { get; }

        /// <summary>
        /// Gets the available server builds.
        /// </summary>
        IList<ITfsServerBuild> GetAvailableServerBuilds();

        /// <summary>
        /// The method returns all available test plans with flag 'Active' and test results for given server build.
        /// </summary>
        /// <param name="serverBuild">Server build to use as filter. If <c>null</c>, parameter is ignored.</param>
        /// <returns>Available test plans.</returns>
        IList<ITfsTestPlan> GetTestPlans(ITfsServerBuild serverBuild);

        /// <summary>
        /// Gets the available test configurations for given test plan.
        /// </summary>
        /// <param name="serverBuild">Server build to use as filter. If <c>null</c>, parameter is ignored.</param>
        /// <param name="testPlan">Test plan to get the available test configurations for</param>
        /// <returns>Available test configurations for given test plan.</returns>
        IList<ITfsTestConfiguration> GetTestConfigurationsForTestPlanWithTestResults(ITfsServerBuild serverBuild, ITfsTestPlan testPlan);

        /// <summary>
        /// Gets the available test configurations for given test plan.
        /// </summary>
        /// <param name="testPlan">Test plan to get the available test configurations for</param>
        /// <returns>Available test configurations for given test plan.</returns>
        IList<ITfsTestConfiguration> GetAllTestConfigurationsForTestPlan(ITfsTestPlan testPlan);

        /// <summary>
        /// Gets a signle <see cref="IWorkItem"/> using the work item ID and the current adapter.
        /// </summary>
        /// <param name="id">The work item id.</param>
        /// <returns></returns>
        IWorkItem GetTfsWorkItemById(int id);

        /*
                /// <summary>
                /// Gets the available test configuratons for the given testplan 
                /// </summary>
                /// <param name="testPlan"></param>
                /// <param name="latest">Option to get only the configuration of the latest test result</param>
                /// <returns></returns>
                IList<ITfsTestConfiguration> GetAssignedTestConfigurationsForTestPlan(ITfsTestPlan testPlan, bool latest, bool allConfigurations);
        */

        /// <summary>
        /// Gets available test configurations for given testcases
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="testCases"></param>
        /// <param name="latest">true if only the latest test results should be considered</param>
        /// <param name="allConfigurations">True if all configurations should be returned</param>
        /// <param name="tfsTestSuite">A test suite for which testcases the configurations should be queried</param>
        /// <returns></returns>
        IEnumerable<ITfsTestConfigurationDetail> GetAssignedTestConfigurationsForTestCases(ITfsTestPlan testPlan, IList<ITfsTestCaseDetail> testCases, bool latest, bool allConfigurations, ITfsTestSuite tfsTestSuite);


        /// <summary>
        /// The method finds and returns the required test plan.
        /// </summary>
        /// <param name="testPlan"><see cref="ITfsTestPlan"/> specifies test plan which should be returned.</param>
        /// <returns>Required test plan.</returns>
        ITfsTestPlanDetail GetTestPlanDetail(ITfsTestPlan testPlan);

        /// <summary>
        /// The method finds and returns the required test plan.
        /// </summary>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> specifies child test suite from which is the test plan to determine.</param>
        /// <returns>Required test plan.</returns>
        ITfsTestPlanDetail GetTestPlanDetail(ITfsTestSuite testSuite);

        /// <summary>
        /// The method finds and returns the required test suite.
        /// </summary>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> specifies test suite which should be returned.</param>
        /// <returns>Required test suite.</returns>
        ITfsTestSuiteDetail GetTestSuiteDetail(ITfsTestSuite testSuite);

        /// <summary>
        /// The method finds and returns the required test configuration.
        /// </summary>
        /// <param name="testConfiguration"><see cref="ITfsTestConfiguration"/> specifies test configuration which should be returned.</param>
        /// <returns>Required test configuration.</returns>
        ITfsTestConfigurationDetail GetTestConfigurationDetail(ITfsTestConfiguration testConfiguration);

        /// <summary>
        /// The method determines all test cases that belongs to the given test plan.
        /// </summary>
        /// <param name="testPlan">Test plan whose test cases should be determined.</param>
        /// <returns>Determined test cases of given test plan.</returns>
        IList<ITfsTestCaseDetail> GetTestCases(ITfsTestPlan testPlan, bool expandSharedSteps);

        /// <summary>
        /// The method determines all test cases that belongs to the given test suite and all dependent test suites.
        /// </summary>
        /// <param name="testSuite">Test suite whose test cases should be determined.</param>
        /// <returns>Determined test cases of given test suite.</returns>
        IList<ITfsTestCaseDetail> GetAllTestCases(ITfsTestSuite testSuite, bool expandSharedSteps);

        /// <summary>
        /// The method determines all test cases that belongs to the given test suite.
        /// </summary>
        /// <param name="testSuite">Test suite whose test cases should be determined.</param>
        /// <returns>Determined test cases of given test suite.</returns>
        IList<ITfsTestCaseDetail> GetTestCases(ITfsTestSuite testSuite, bool expandSharedSteps);

        //  <param name="hierarchical">Determines if testresults for an testsuite should be searched hierarchical</param>
        /// <summary>
        /// The method determines all test results that belongs to the given test case.
        /// </summary>
        /// <param name="testPlan">Test plan whose test results should be determined.</param>
        /// <param name="testCase">Test case whose test results should be determined.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        ///     If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        ///     If <c>null</c>, the build is no matter.</param>
        /// <param name="testSuite"></param>
        /// <param name="hierarchical">True if the test results should be queried for the actual hierarchy of test suites</param>
        /// <returns>Determined test results of given test case.</returns>
        IList<ITfsTestResultDetail> GetTestResults(ITfsTestPlan testPlan, ITfsTestCase testCase, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild, ITfsTestSuite testSuite, bool hierarchical, out Dictionary<int, int> lastTestRunPerConfig);

        /// <summary>
        /// The method determines only the latest test results that belongs to the given test case.
        /// </summary>
        /// <param name="testPlan">Test plan whose test results should be determined.</param>
        /// <param name="testCase">Test case whose test results should be determined.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        ///     If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        ///     If <c>null</c>, the build is no matter.</param>
        /// <param name="testSuite"></param>
        /// <param name="hierarchical">Determines if testresults for an testsuite should be searched hierarchical</param>
        /// <returns>Determined test results of given test case.</returns>
        IList<ITfsTestResultDetail> GetLatestTestResults(ITfsTestPlan testPlan, ITfsTestCase testCase, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild, ITfsTestSuite testSuite, bool hierarchical);


        /// <summary>
        /// The method determines the latest test results that belongs to the given test case for all provided configurations
        /// </summary>
        /// <param name="testPlan">Test plan whose test results should be determined.</param>
        /// <param name="testCase">Test case whose test results should be determined.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        ///     If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        ///     If <c>null</c>, the build is no matter.</param>
        /// <param name="testSuite"></param>
        ///   /// <param name="hierarchical">Determines if testresults for an testsuite should be searched hierarchical</param>
        /// <returns>Determined test results of given test case.</returns>
        IList<ITfsTestResultDetail> GetLatestTestResultsForAllSelectedConfigurations(ITfsTestPlan testPlan, ITfsTestCase testCase, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild, ITfsTestSuite testSuite, bool hierarchical);


        /// <summary>
        /// The method determines if at least one test result exists for given test case.
        /// </summary>
        /// <param name="testPlan">Test plan whose test results should be determined.</param>
        /// <param name="testCase">Test case to determine if at least one test result exists.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        /// If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        /// If <c>null</c>, the build is no matter.</param>
        /// <returns><c>true</c> for the test cases exists at least one test result. Otherwise <c>false</c>.</returns>
        bool TestResultExists(ITfsTestPlan testPlan, ITfsTestSuite testSuite, ITfsTestCaseDetail testCase, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild);

        /// <summary>
        /// The method determines if at least one test result exists for given test cases.
        /// </summary>
        /// <param name="testPlan">Test plan whose test results should be determined.</param>
        /// <param name="testCases">Test cases to determine if at least one test result exists.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        /// If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        /// If <c>null</c>, the build is no matter.</param>
        /// <returns><c>true</c> for the test cases exists at least one test result. Otherwise <c>false</c>.</returns>
        bool TestResultExists(ITfsTestPlan testPlan, ITfsTestSuite testSuite, IEnumerable<ITfsTestCaseDetail> testCases, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild);

        /// <summary>
        /// The method determines the hyperlink for the build to the team foundation server web access.
        /// </summary>
        /// <param name="buildNumber">Build number of the build.</param>
        /// <returns>Required hyperlink.</returns>
        Uri GetBuildViewerLink(string buildNumber);

        /// <summary>
        /// The method determines the hyperlink for the work item to the team foundation server web access editor of the work item.
        /// </summary>
        /// <param name="workItemArtifactLink">Artifact link of the work item.</param>
        /// <returns>Required hyperlink.</returns>
        Uri GetWorkItemEditorLink(string workItemArtifactLink);

        /// <summary>
        /// The method determines the hyperlink for the work item to the team foundation server web access viewer of the work item.
        /// </summary>
        /// <param name="workItemNumber">Number of the work item.</param>
        /// <param name="revision">Required revision of work item. If not defined, last revision is used.</param>
        /// <returns>Required hyperlink.</returns>
        Uri GetWorkItemViewerLink(int workItemNumber, int revision);

        /// <summary>
        /// Gets the artifact link.
        /// Valid artifact types are: <code>Project, WorkItem, Query, VersionedItem, LatestItemVersion, Changeset, Shelveset, ShelvedItem and Build</code>.
        /// </summary>
        /// <param name="artifactLink">A valid team foundation server artifact Uri.</param>
        /// <returns>A viewer URL for a TFS artifact.</returns>
        Uri GetArtifactLink(string artifactLink);

        /// <summary>
        /// Gets the title of work item.
        /// </summary>
        /// <param name="workItemId">Id of the work item to get the title for.</param>
        /// <returns>Title of the work item if the work item found. Otherwise <c>null</c>.</returns>
        string GetWorkItemTitle(int workItemId);

        /// <summary>
        /// The method expands the enumerable. Used at least for shared steps in test case actions.
        /// </summary>
        /// <param name="enumerable">Enumerable to expand.</param>
        /// <returns>Expanded enumerable.</returns>
        IEnumerable ExpandEnumerable(IEnumerable enumerable);

        /// <summary>
        /// Method to get all LinkedWorkItems For A given WorkItemId that meet special requirments.
        /// This method can be used if the "Links" property is not available i.e. TestResult with attached defects
        /// </summary>
        /// <param name="workItemId">the Id of the WorkItem</param>
        /// <param name="workItemTypeCsv">The Type of the linked work item </param>
        /// <param name="linkTypeCsv">the link type</param>
        /// <param name="filterOption"></param>
        /// <returns></returns>
        List<WorkItem> GetAllLinkedWorkItemsForWorkItemId(int workItemId, string workItemTypeCsv, string linkTypeCsv, IFilterOption filterOption);

        /// <summary>
        /// This method gets all linked workitems for a test case
        /// </summary>
        /// <param name="testCases"></param>
        /// <param name="workItemTypes"></param>
        /// <param name="linkFilter"></param>
        /// <param name="filterOption"></param>
        /// <returns></returns>
        List<WorkItem> GetAllLinkedWorkItemsForTestCases(IList<ITfsTestCaseDetail> testCases, List<string> workItemTypes, IWorkItemLinkFilter linkFilter, IFilterOption filterOption);


        /// <summary>
        /// This method gets all linked workitems for a test result
        /// </summary>
        /// <param name="testResult">The test result</param>
        /// <param name="workItemTypes">The work item types that should be returned</param>
        /// <param name="linkFilter">the link filters</param>
        /// <param name="filterOption">the filter options</param>
        /// <returns></returns>
        List<WorkItem> GetAllLinkedWorkItemsForTestResult(ITfsTestResultDetail testResult, List<string> workItemTypes, IWorkItemLinkFilter linkFilter, IFilterOption filterOption);


        /// <summary>
        /// This method queries all builds for a testcase
        /// </summary>
        /// <param name="tfsTestCaseDetail"></param>
        /// <returns></returns>
        IEnumerable<IBuildDetail> GetAllBuildsForTestCase(ITfsTestCaseDetail tfsTestCaseDetail, bool reloadBuilds);

        /// <summary>
        /// This method queries builds for a specific test case
        /// </summary>
        /// <param name="tfsTestCaseDetail">The specific Test Case</param>
        /// <param name="latestTestResults">Filter for the opertation, if true only the builds for the latest test result are returned</param>
        /// <param name="latestTestResultsForAllConfigurations">Filter for the opertation, if true only the builds for the latest test result for all configurations are returned</param>
        /// <param name="testConfiguration">The test configuration that should bs used</param>
        /// <returns></returns>
        IEnumerable<IBuildDetail> GetBuildsForTestCase(ITfsTestCaseDetail tfsTestCaseDetail, bool latestTestResults, bool latestTestResultsForAllConfigurations, ITfsTestConfiguration testConfiguration, bool reloadBuilds);


        /// <summary>
        /// This method gets the test Points for a TestCase
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="testCases"></param>
        /// <returns></returns>
        IEnumerable<ITfsPropertyValueProvider> GetTestPointsForTestCases(ITfsTestPlan testPlan, IList<ITfsTestCaseDetail> testCases);

        /// <summary>
        /// Get the test points for a given testplan
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="boundElement"></param>
        /// <returns></returns>
        IList<ITfsTestPointDetail> GetTestPointsForTestPlan(ITfsTestPlan testPlan, ITfsPropertyValueProvider boundElement);

        /// <summary>
        /// Get the test points for a given Test Suite
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="boundElement"></param>
        /// <returns></returns>
        IList<ITfsTestPointDetail> GetTestPointsForTestSuite(ITfsTestPlan testPlan, ITfsPropertyValueProvider boundElement);
    }
}
