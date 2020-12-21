using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Enums.Model;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// Configuration class to configure default values for specification report pane.
    /// </summary>
    public class TestSpecRepDefaultValueXml : ITestSpecReportDefault
    {
        [XmlElement("SelectTestPlan")]
        public string SelectTestPlan { get; set; }


        [XmlElement("SelectTestSuite")]
        public string SelectTestSuite { get; set; }

        [XmlElement("CreateDocumentStructure")]
        public bool CreateDocumentStructure { get; set; }

        [XmlElement("DocumentStructureType")]
        public DocumentStructureType DocumentStructureType { get; set; }

        [XmlElement("SkipLevels")]
        public int SkipLevels { get; set; }

        [XmlElement("IncludeTestConfigurations")]
        public bool IncludeTestConfigurations { get; set; }

        [XmlElement("ConfigurationPositionType")]
        public ConfigurationPositionType ConfigurationPositionType { get; set; }

        [XmlElement("SortTestCasesBy")]
        public TestCaseSortType TestCaseSortType { get; set; }
    }
}
