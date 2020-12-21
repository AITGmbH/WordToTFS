namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    public interface ITestSuiteSearcher
    {
        /// <summary>
        /// Searches for a test suite within the test suite hiearchy below the test plan.
        /// </summary>
        /// <param name="suiteNameOrPath">The sute name or path wihch should be found.</param>
        /// <returns>The found test suite.</returns>
        ITfsTestSuite SearchTestSuiteWithinTestPlan(string suiteNameOrPath);
    }
}
