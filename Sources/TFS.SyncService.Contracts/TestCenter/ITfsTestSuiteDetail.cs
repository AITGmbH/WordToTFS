namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test suite.
    /// </summary>
    public interface ITfsTestSuiteDetail : ITfsPropertyValueProvider
    {
        /// <summary>
        /// Gets the base information.
        /// </summary>
        ITfsTestSuite TestSuite { get; }
    }
}
