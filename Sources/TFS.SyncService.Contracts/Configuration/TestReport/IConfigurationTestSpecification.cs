using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for test specification report area.
    /// </summary>
    public interface IConfigurationTestSpecification
    {
        /// <summary>
        /// Gets flag saying whether is the test specification available or not.
        /// </summary>
        bool Available { get; }

        /// <summary>
        /// Gets the template file for test plan.
        /// </summary>
        string TestPlanTemplate { get; }

        /// <summary>
        /// Gets the template file for test suite.
        /// </summary>
        string TestSuiteTemplate { get; }

        /// <summary>
        /// Gets the template file for the root test suite
        /// </summary>
        string RootTestSuiteTemplate { get; }

        /// <summary>
        /// Gets the template file for the root test suite
        /// </summary>
        string LeafTestSuiteTemplate { get; }

        /// <summary>
        /// Gets the template file for test case element.
        /// </summary>
        string TestCaseElementTemplate { get; }

        /// <summary>
        /// Gets the template file for shared steps element
        /// </summary>
        string SharedStepsElementTemplate { get; }

        /// <summary>
        /// A summary page that can be inserted at the end of the report
        /// </summary>
        string SummaryPageTemplate { get; }

        /// <summary>
        /// Gets the list of pre-operations.
        /// </summary>
        IList<IConfigurationTestOperation> PreOperations { get; }

        /// <summary>
        /// Gets the list of post-operations.
        /// </summary>
        IList<IConfigurationTestOperation> PostOperations { get; }

        ITestSpecReportDefault DefaultValues { get; }

        /// <summary>
        /// The template to display the configurations
        /// </summary>
        string TestConfigurationElementTemplate
        {
            get;
            
        }
    }
}
