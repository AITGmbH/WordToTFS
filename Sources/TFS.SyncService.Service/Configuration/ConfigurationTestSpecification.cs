using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationTestSpecification"/> - configuration for test specification report area.
    /// </summary>
    public class ConfigurationTestSpecification : IConfigurationTestSpecification
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestSpecification"/> class.
        /// </summary>
        public ConfigurationTestSpecification()
        {
            Available = false;
            TestPlanTemplate = null;
            TestSuiteTemplate = null;
            RootTestSuiteTemplate = null;
            LeafTestSuiteTemplate = null;
            TestCaseHeaderTemplate = null;
            TestCaseElementTemplate = null;
            SharedStepsElementTemplate = null;
            TestConfigurationElementTemplate = null;
            SummaryPageTemplate = null;
            PreOperations = new List<IConfigurationTestOperation>();
            PostOperations = new List<IConfigurationTestOperation>();
            DefaultValues = null;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestSpecification"/> class.
        /// </summary>
        /// <param name="testSpecificationConfiguration">Associated test specification area configuration from w2t file.</param>
        public ConfigurationTestSpecification(TestSpecificationConfiguration testSpecificationConfiguration)
            : this()
        {
            TestSpecificationConfiguration = testSpecificationConfiguration;
            if (TestSpecificationConfiguration != null)
            {
                Available = TestSpecificationConfiguration.Available;
                TestPlanTemplate = TestSpecificationConfiguration.TestPlanTemplate;
                TestSuiteTemplate = TestSpecificationConfiguration.TestSuiteTemplate;
                RootTestSuiteTemplate = TestSpecificationConfiguration.RootTestSuiteTemplate ?? TestSuiteTemplate;
                LeafTestSuiteTemplate = TestSpecificationConfiguration.LeafTestSuiteTemplate ?? TestSuiteTemplate;
                TestCaseElementTemplate = TestSpecificationConfiguration.TestCaseElementTemplate;
                SharedStepsElementTemplate = TestSpecificationConfiguration.SharedStepsElementTemplate;
                TestConfigurationElementTemplate = testSpecificationConfiguration.TestConfigurationElementTemplate;
                SummaryPageTemplate = TestSpecificationConfiguration.SummaryPageTemplate;
                DefaultValues = testSpecificationConfiguration.DefaultValues;

                if (TestSpecificationConfiguration.PreOperations != null && TestSpecificationConfiguration.PreOperations.Count > 0)
                {
                    PreOperations = new List<IConfigurationTestOperation>();
                    foreach (var operation in TestSpecificationConfiguration.PreOperations)
                        PreOperations.Add(new ConfigurationTestOperation(operation));
                }
                if (TestSpecificationConfiguration.PostOperations != null && TestSpecificationConfiguration.PostOperations.Count > 0)
                {
                    PostOperations = new List<IConfigurationTestOperation>();
                    foreach (var operation in TestSpecificationConfiguration.PostOperations)
                        PostOperations.Add(new ConfigurationTestOperation(operation));
                }
            }
        }

        #endregion Constructors

        #region Internal properties

        /// <summary>
        /// Gets the test specification area configuration read from w2t file. Set only in one constructor.
        /// </summary>
        internal TestSpecificationConfiguration TestSpecificationConfiguration { get; private set; }

        #endregion Internal properties

        #region Implementation of IConfigurationTestSpecification

        /// <summary>
        /// Gets flag saying whether is the test specification available or not.
        /// </summary>
        public bool Available { get; private set; }

        /// <summary>
        /// Gets the template file for test plan.
        /// </summary>
        public string TestPlanTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test suite.
        /// </summary>
        public string TestSuiteTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for the root test suite
        /// </summary>
        public string RootTestSuiteTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for the root test suite
        /// </summary>
        public string LeafTestSuiteTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test case header.
        /// </summary>
        public string TestCaseHeaderTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for test case element.
        /// </summary>
        public string TestCaseElementTemplate { get; private set; }

        /// <summary>
        /// Gets the template file for shared steps element
        /// </summary>
        public string SharedStepsElementTemplate { get; private set; }

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
        public IList<IConfigurationTestOperation> PreOperations
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the list of post-operations.
        /// </summary>
        public IList<IConfigurationTestOperation> PostOperations
        {
            get;
            private set;
        }

        public ITestSpecReportDefault DefaultValues
        {
            get;
            private set;
        }


        /// <summary>
        /// The template to display the configurations
        /// </summary>
        public string TestConfigurationElementTemplate
        {
            get;
            private set;
        }

        #endregion Implementation of IConfigurationTestSpecification
    }
}
