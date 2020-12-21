using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationTestResult"/> - configuration for test result report area.
    /// </summary>
    public class ConfigurationTestResult : IConfigurationTestResult
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestResult"/> class.
        /// </summary>
        public ConfigurationTestResult()
        {
            Available = false;
            IncludeTestCasesWithoutResults = false;
            TestPlanTemplate = null;
            TestSuiteTemplate = null;
            RootTestSuiteTemplate = null;
            LeafTestSuiteTemplate = null;
            TestCaseHeaderTemplate = null;
            TestCaseElementTemplate = null;
            TestResultHeaderTemplate = null;
            TestResultElementTemplate = null;
            TestConfigurationHeaderTemplate = null;
            TestConfigurationElementTemplate = null;
            SummaryPageTemplate = null;
            BuildQualities = new List<string>();
            PreOperations = new List<IConfigurationTestOperation>();
            PostOperations = new List<IConfigurationTestOperation>();
            DefaultValues = null;
            BuildFilters = null;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestResult"/> class.
        /// </summary>
        /// <param name="testResultConfiguration">Associated test result area configuration from w2t file.</param>
        public ConfigurationTestResult(TestResultConfiguration testResultConfiguration)
            :this()
        {
            TestResultConfiguration = testResultConfiguration;
            if (TestResultConfiguration == null)
                return;

            Available = TestResultConfiguration.Available;
            IncludeTestCasesWithoutResults = TestResultConfiguration.IncludeTestCasesWithoutResults;
            TestPlanTemplate = TestResultConfiguration.TestPlanTemplate;
            TestSuiteTemplate = TestResultConfiguration.TestSuiteTemplate;
            RootTestSuiteTemplate = TestResultConfiguration.RootTestSuiteTemplate ?? TestSuiteTemplate;
            LeafTestSuiteTemplate = TestResultConfiguration.LeafTestSuiteTemplate ?? TestSuiteTemplate;
            TestCaseElementTemplate = TestResultConfiguration.TestCaseElementTemplate;
            TestResultElementTemplate = TestResultConfiguration.TestResultElementTemplate;
            TestConfigurationElementTemplate = TestResultConfiguration.TestConfigurationElementTemplate;
            SummaryPageTemplate = TestResultConfiguration.SummaryPageTemplate;
            BuildQualities = TestResultConfiguration.BuildQualities;
            BuildFilters = TestResultConfiguration.BuildFilters;
            DefaultValues = TestResultConfiguration.DefaultValues;

            if (TestResultConfiguration.PreOperations != null && TestResultConfiguration.PreOperations.Count > 0)
            {
                PreOperations = new List<IConfigurationTestOperation>();
                foreach (var operation in TestResultConfiguration.PreOperations)
                    PreOperations.Add(new ConfigurationTestOperation(operation));
            }
            if (TestResultConfiguration.PostOperations != null && TestResultConfiguration.PostOperations.Count > 0)
            {
                PostOperations = new List<IConfigurationTestOperation>();
                foreach (var operation in TestResultConfiguration.PostOperations)
                    PostOperations.Add(new ConfigurationTestOperation(operation));
            }
        }

        #endregion Constructors

        #region Internal properties

        /// <summary>
        /// Gets the test result area configuration read from w2t file. Set only in one constructor.
        /// </summary>
        internal TestResultConfiguration TestResultConfiguration { get; private set; }

        #endregion Internal properties

        #region Implementation of IConfigurationTestResult

        /// <summary>
        /// Gets flag saying whether is the test result available or not.
        /// </summary>
        public bool Available { get; private set; }

        /// <summary>
        /// Gets flag saying whether the test case without result is included
        /// </summary>
        public bool IncludeTestCasesWithoutResults  { get; private set; }

        /// <summary>
        /// Gets the template file for test plan
        /// </summary>
        public string TestPlanTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test suite
        /// </summary>
        public string TestSuiteTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for the root test suite
        /// </summary>
        public string RootTestSuiteTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for the leaf test suite
        /// </summary>
        public string LeafTestSuiteTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test case header
        /// </summary>
        public string TestCaseHeaderTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test case element
        /// </summary>
        public string TestCaseElementTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for shared steps element
        /// </summary>
        public string SharedStepsElementTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test result header
        /// </summary>
        public string TestResultHeaderTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test result element
        /// </summary>
        public string TestResultElementTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test configuration header
        /// </summary>
        public string TestConfigurationHeaderTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test configuration element
        /// </summary>
        public string TestConfigurationElementTemplate { get; private set; }

        /// <summary>
        /// Gets the collection of configured build qualities to use as filter for used builds.
        /// </summary>
        public IList<string> BuildQualities { get; private set; }

        /// <summary>
        /// Gets the collection of configured build filters to use as filter for used builds.
        /// </summary>
        public IBuildFilters BuildFilters { get; private set; }

        /// <summary>
        /// A summary page that can be inserted at the end of the report
        /// </summary>
        public string SummaryPageTemplate
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the list of pre-operations.
        /// </summary>
       public IList<IConfigurationTestOperation> PreOperations { get;
            private set;
       }

        /// <summary>
        /// Gets the list of post-operations.
        /// </summary>
        public IList<IConfigurationTestOperation> PostOperations { get;
            private set;
        }

        public ITestResultReportDefault DefaultValues { get; private set; }

        #endregion Implementation of IConfigurationTestResult
    }
}
