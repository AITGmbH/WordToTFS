namespace AIT.TFS.SyncService.Contracts.Adapter
{
    using WorkItemObjects;
    using Configuration;

    /// <summary>
    /// Interface for team foundation server adapter factory.
    /// </summary>
    public interface ITfsAdapterCreator
    {
        /// <summary>
        /// Creates an instance of team foundation server adapter.
        /// </summary>
        /// <param name="serverName">Name of the TFS server to work with.</param>
        /// <param name="projectName">Name of the TFS project to work with.</param>
        /// <param name="queryConfiguration">A <see cref="IQueryConfiguration"/> describing what work items to load</param>
        /// <param name="configuration">Work item mapping configuration</param>
        /// <returns>Created adapter.</returns>
        IWorkItemSyncAdapter CreateAdapter(string serverName, string projectName, IQueryConfiguration queryConfiguration, IConfiguration configuration);
    }
}