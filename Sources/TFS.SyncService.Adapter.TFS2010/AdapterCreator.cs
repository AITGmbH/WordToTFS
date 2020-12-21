namespace AIT.TFS.SyncService.Adapter.TFS2012
{
    using Contracts.Adapter;
    using Contracts.WorkItemObjects;
    using Contracts.Configuration;

    /// <summary>
    /// Adapter factory for team foundation server 2012
    /// </summary>
    internal class AdapterCreator : ITfsAdapterCreator
    {
        #region ITfsAdapterCreator Members

        /// <summary>
        /// Creates an instance of team foundation server adapter.
        /// </summary>
        /// <param name="serverName">Name of the TFS server to work with.</param>
        /// <param name="projectName">Name of the TFS project to work with.</param>
        /// <param name="queryConfiguration">A <see cref="IQueryConfiguration"/> describing what work items to load</param>
        /// <param name="configuration">Work item mapping configuration</param>
        /// <returns>Created adapter.</returns>
        public IWorkItemSyncAdapter CreateAdapter(string serverName, string projectName, IQueryConfiguration queryConfiguration, IConfiguration configuration)
        {
            return new Tfs2012SyncAdapter(serverName, projectName, queryConfiguration, configuration);
        }

        #endregion
    }
}