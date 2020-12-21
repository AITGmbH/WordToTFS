namespace AIT.TFS.SyncService.Contracts.ProgressService
{
    /// <summary>
    /// Interface to check if reporting generation should be continued.
    /// </summary>
    public interface ITestReportingProgressCancellationService
    {
        /// <summary>
        /// Returns <c>true </c> if reporting gereration steps should be continued or false if further reporting generation steps should be cancelled.
        /// </summary>
        /// <returns></returns>
        bool CheckIfContinue();
    }
}
