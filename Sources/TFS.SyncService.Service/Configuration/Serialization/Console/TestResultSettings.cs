using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.Console
{
    /// <summary>
    /// Base class of test specification configuration.
    /// </summary>
    [XmlRoot("TestResult", Namespace = "", IsNullable = false)]
    public class TestResultSettings
    {
        [XmlAttribute("Build")]
        public string Build { get; set; }

        [XmlAttribute("TestConfiguration")]
        public string TestConfiguration { get; set; }

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

        [XmlAttribute("IncludeOnlyMostRecentResults")]
        public bool IncludeOnlyMostRecentResults { get; set; }

        [XmlAttribute("MostRecentForAllConfigurations")]
        public bool MostRecentForAllConfigurations { get; set; }
    }
}
