using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Enums.Model;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// Configuration class to configure default values for specification report pane.
    /// </summary>
    public class TestResultRepDefaultValueXml : ITestResultReportDefault
    {

        [XmlElement("CreateDocumentStructure")]
        public bool CreateDocumentStructure { get; set; }

        [XmlElement("DocumentStructureType")]
        public DocumentStructureType DocumentStructureType { get; set; }

        [XmlElement("SelectBuild")]
        public string SelectBuild { get; set; }

        [XmlElement("SelectTestConfiguration")]
        public string SelectTestConfiguration { get; set; }

        [XmlElement("SelectTestPlan")]
        public string SelectTestPlan { get; set; }

        [XmlElement("SkipLevels")]
        public int SkipLevels { get; set; }

        [XmlElement("IncludeTestConfigurations")]
        public bool IncludeTestConfigurations { get; set; }

        [XmlElement("ConfigurationPositionType")]
        public ConfigurationPositionType ConfigurationPositionType { get; set; }

        [XmlElement("SortTestCasesBy")]
        public TestCaseSortType TestCaseSortType { get; set; }

        [XmlElement("IncludeMostRecentTestResult")]
        public bool IncludeMostRecentTestResult { get; set; }

        [XmlElement("IncludeMostRecentTestResultForAllSelectedConfigurations")]
        public bool IncludeMostRecentTestResultForAllSelectedConfigurations { get; set; }
    }
}
