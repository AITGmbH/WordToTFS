using System.Collections.Generic;
using System.ComponentModel;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
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
            TestSpecificationConfiguration = new TestSpecificationConfiguration();
            TestResultConfiguration = new TestResultConfiguration();
            TemplatesConfiguration = new List<TemplateConfiguration>();
        }

        #endregion Constructors

        #region Public serializable properties

        /// <summary>
        /// Gets or sets a value indicating whether the plugin sets the hyperlink base of the document to the document root in order for relative attachment links to work.
        /// </summary>
        [XmlAttribute("SetHyperlinkBase")]
        [DefaultValue(false)]
        public bool SetHyperlinkBase { get; set; }

        /// <summary>
        /// Gets a value indicating whether to expand shared steps for this template.
        /// </summary>
        /// <value>
        ///   <c>true</c> if [shared steps]; otherwise, <c>false</c>.
        /// </value>
        [XmlAttribute("ExpandSharedSteps")]
        public bool ExpandSharedSteps { get; set; }

        /// <summary>
        /// Gets or sets the configuration for test specification.
        /// </summary>
        [XmlElement("TestSpecificationConfiguration", IsNullable = false)]
        public TestSpecificationConfiguration TestSpecificationConfiguration { get; set; }

        /// <summary>
        /// Gets or sets the configuration for test specification.
        /// </summary>
        [XmlElement("TestResultConfiguration", IsNullable = false)]
        public TestResultConfiguration TestResultConfiguration { get; set; }

        /// <summary>
        /// Gets or sets the configuration for templates.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
        [XmlArray("Templates")]
        [XmlArrayItem("Template")]
        public List<TemplateConfiguration> TemplatesConfiguration { get; set; }

        #endregion Public serializable properties
  }
}
