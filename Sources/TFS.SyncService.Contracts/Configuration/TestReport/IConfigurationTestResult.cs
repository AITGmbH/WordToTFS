using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for test result report area.
    /// </summary>
    public interface IConfigurationTestResult
    {
        /// <summary>
        /// Gets flag saying whether is the test specification available or not.
        /// </summary>
        bool Available { get; }

        /// <summary>
        /// Gets flag saying whether the test case without result is included
        /// </summary>
        bool IncludeTestCasesWithoutResults { get; }

        /// <summary>
        /// Gets the template file for test plan
        /// </summary>
        string TestPlanTemplate { get; }

        /// <summary>
        /// Gets the template file for test suite
        /// </summary>
        string TestSuiteTemplate { get; }

        /// <summary>
        /// Gets the template file for the root test suite
        /// </summary>
        string RootTestSuiteTemplate { get; }

        /// <summary>
        /// Gets the template file for the leaf test suite
        /// </summary>
        string LeafTestSuiteTemplate { get; }

        /// <summary>
        /// Gets the template file for test case element
        /// </summary>
        string TestCaseElementTemplate { get; }

        /// <summary>
        /// Gets the template file for shared steps element
        /// </summary>
        string SharedStepsElementTemplate { get; }

        /// <summary>
        /// Gets the template file for test result element
        /// </summary>
        string TestResultElementTemplate { get; }

        /// <summary>
        /// Gets the template file for test configuration element
        /// </summary>
        string TestConfigurationElementTemplate { get; }

        /// <summary>
        /// Gets the collection of configured build qualities to use as filter for used builds.
        /// </summary>
        IList<string> BuildQualities { get; }

        /// <summary>
        /// Gets the collection of configured build filters to use as filter for used builds.
        /// </summary>
        IBuildFilters BuildFilters { get; }

        /// <summary>
        /// A summary page that can be inserted at the end of the report
        /// </summary>
        string SummaryPageTemplate { get; }

        /// <summary>
        /// PreOperations for the TestResultReport
        /// </summary>
        IList<IConfigurationTestOperation> PreOperations
        {
            get;

        }
        /// <summary>
        /// PostOperations for the TestResultReport
        /// </summary>
        IList<IConfigurationTestOperation> PostOperations
        {
            get;
        }

        ITestResultReportDefault DefaultValues { get; }
    }
}
