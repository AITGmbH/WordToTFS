using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// Base class of test specification configuration.
    /// </summary>
    [XmlRoot("TestResultConfiguration")]
    public class TestResultConfiguration
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TestResultConfiguration"/> class.
        /// </summary>
        public TestResultConfiguration()
        {
            Available = false;
        }

        #endregion Constructors

        #region Public serialization properties

        /// <summary>
        /// Gets or sets flag saying whether is the test specification available or not.
        /// </summary>
        [XmlAttribute("Available")]
        public bool Available { get; set; }

        /// <summary>
        /// Gets flag saying whether the test case without result is included
        /// </summary>
        [XmlAttribute("IncludeTestCasesWithoutResults")]
        public bool IncludeTestCasesWithoutResults { get; set; }

        /// <summary>
        /// Gets or sets the template file for test plan
        /// </summary>
        [XmlAttribute("TestPlanTemplate")]
        public string TestPlanTemplate { get; set; }

        /// <summary>
        /// Gets or sets the template file for test suite
        /// </summary>
        [XmlAttribute("TestSuiteTemplate")]
        public string TestSuiteTemplate { get; set; }

        /// <summary>
        /// Gets or sets the template file for the root test suite
        /// </summary>
        [XmlAttribute("RootTestSuiteTemplate")]
        public string RootTestSuiteTemplate { get; set; }

        /// <summary>
        /// Gets or sets the template file for the leaf test suite
        /// </summary>
        [XmlAttribute("LeafTestSuiteTemplate")]
        public string LeafTestSuiteTemplate { get; set; }

        /// <summary>
        /// Gets or sets the template file for test case element
        /// </summary>
        [XmlAttribute("TestCaseElementTemplate")]
        public string TestCaseElementTemplate { get; set; }

        /// <summary>
        /// Gets or sets the template file for shared steps element
        /// </summary>
        [XmlAttribute("SharedStepsElementTemplate")]
        public string SharedStepsElementTemplate { get; set; }

        /// <summary>
        /// Gets or sets the template file for test result element
        /// </summary>
        [XmlAttribute("TestResultElementTemplate")]
        public string TestResultElementTemplate { get; set; }

        /// <summary>
        /// Gets or sets the template file for configuration element
        /// </summary>
        [XmlAttribute("TestConfigurationElementTemplate")]
        public string TestConfigurationElementTemplate { get; set; }

        /// <summary>
        /// Gets the list of configured build qualities to use as filter for used builds.
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
        [XmlArray("BuildQualities")]
        [XmlArrayItem("BuildQuality")]
        public List<string> BuildQualities { get; set; }


        /// <summary>
        /// Gets the collection of configured build filters to use as filter for used builds.
        /// </summary>
        [XmlElement("BuildFilters")]
        public BuildFilters BuildFilters { get; set; }

        /// <summary>
        /// A summary page that can be inserted at the end of the report
        /// </summary>
        [XmlAttribute("SummaryPageTemplate")]
        public string SummaryPageTemplate
        {
            get;
            set;
        }

         /// <summary>
         /// Property used to serialize pre-operations.
         /// </summary>
         [SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
         [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
         [XmlArray("PreOperations")]
         [XmlArrayItem("Operation")]
         public List<OperationConfiguration> PreOperations { get; set; }

         /// <summary>
         /// Property used to serialize post-operations
         /// </summary>
         [SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
         [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
         [XmlArray("PostOperations")]
         [XmlArrayItem("Operation")]
         public List<OperationConfiguration> PostOperations { get; set; }

         [XmlElement("DefaultValues")]
         public TestResultRepDefaultValueXml DefaultValues { get; set; }


        #endregion Public serialization properties
    }
}
