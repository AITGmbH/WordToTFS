using AIT.TFS.SyncService.Contracts.Enums.Model;

namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    public interface ITestResultReportDefault
    {
        string SelectBuild { get; set; }

        string SelectTestConfiguration { get; set; }

        string SelectTestPlan { get; set; }

        int SkipLevels { get; set; }

        bool CreateDocumentStructure { get; set; }

        DocumentStructureType DocumentStructureType { get; set; }

        bool IncludeTestConfigurations { get; set; }

        ConfigurationPositionType ConfigurationPositionType { get; set; }

        TestCaseSortType TestCaseSortType { get; set; }

        bool IncludeMostRecentTestResult { get; set; }

        bool IncludeMostRecentTestResultForAllSelectedConfigurations { get; set; }

    }
}