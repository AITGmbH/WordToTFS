namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test configuration.
    /// </summary>
    public interface ITfsTestConfigurationDetail : ITfsPropertyValueProvider
    {
        /// <summary>
        /// Gets the base information.
        /// </summary>
        ITfsTestConfiguration TestConfiguration { get; }
    }
}
