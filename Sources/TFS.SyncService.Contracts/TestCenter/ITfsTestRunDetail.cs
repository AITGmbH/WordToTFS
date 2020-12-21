namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test run.
    /// </summary>
    public interface ITfsTestRunDetail : ITfsPropertyValueProvider
    {
        /// <summary>
        /// Gets the base information.
        /// </summary>
        ITfsTestRun TestRun { get; }
    }
}
