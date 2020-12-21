using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.Console
{
    /// <summary>
    /// Base class of test configuration.
    /// </summary>
    [XmlRoot("TestConfiguration")]
    public class TestConfiguration
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TestConfiguration"/> class.
        /// </summary>
        public TestConfiguration()
        {
            TestSpecificationConfiguration = new TestSpecSettings();
            
        }

        #endregion Constructors

        #region Public serializable properties

        /// <summary>
        /// Gets or sets the configuration for test specification report.
        /// </summary>
        [XmlElement("TestSpecificationConfiguration")]
        public TestSpecSettings TestSpecificationConfiguration { get; set; }

        /// <summary>
        /// Gets or sets the configuration for test result report.
        /// </summary>
        [XmlElement("TestResultConfiguration")]
        public TestResultSettings TestResultConfiguration { get; set; }

        #endregion Public serializable properties
    }
}
