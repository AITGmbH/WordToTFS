namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test run.
    /// </summary>
    public interface ITfsTestRun
    {
        /// <summary>
        /// Gets id of the test run.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets id of test settings.
        /// </summary>
        int TestSettingsId { get; }

        /// <summary>
        /// Gets title of the test run.
        /// </summary>
        string Title { get; }
    }
}
