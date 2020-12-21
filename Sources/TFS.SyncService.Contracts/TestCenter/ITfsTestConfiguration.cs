namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test configuration.
    /// </summary>
    public interface ITfsTestConfiguration
    {
        /// <summary>
        /// Gets the id of the test configuration.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets the name of test configuration.
        /// </summary>
        string Name { get; }
    }
}
