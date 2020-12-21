using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.Console
{
    /// <summary>
    /// Base class of test specification configuration.
    /// </summary>
    [XmlRoot("TestSpecification", Namespace = "", IsNullable = false)]
    public class TestSpecSettings
    {
        [XmlAttribute("TestPlan")]
        public string TestPlan { get; set; }

        [XmlAttribute("TestSuite")]
        public string TestSuite { get; set; }

        [XmlAttribute("CreateDocumentStructure")]
        public bool CreateDocumentStructure { get; set; }

        [XmlAttribute("DocumentStructure")]
        public string DocumentStructure { get; set; }

        [XmlAttribute("SkipLevels")]
        public int SkipLevels { get; set; }

        [XmlAttribute("IncludeTestConfigurations")]
        public bool IncludeTestConfigurations { get; set; }

        [XmlAttribute("TestConfigurationsPosition")]
        public string TestConfigurationsPosition { get; set; }

        [XmlAttribute("SortTestCasesBy")]
        public string SortTestCasesBy { get; set; }
    }
}
