using AIT.TFS.SyncService.Contracts.Enums.Model;

namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    public interface ITestSpecReportDefault
    {
        string SelectTestPlan { get; set; }

        string SelectTestSuite { get; set; }

        bool CreateDocumentStructure { get; set; }

        DocumentStructureType DocumentStructureType { get; set; }

        int SkipLevels { get; set; }

        bool IncludeTestConfigurations { get; set; }

        ConfigurationPositionType ConfigurationPositionType { get; set; }

        TestCaseSortType TestCaseSortType { get; set; }

    }
}