namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemCollections
{
    #region Usings
    using System.Collections.ObjectModel;
    using System.Linq;
    using Contracts.WorkItemCollections;
    using Contracts.WorkItemObjects;
    #endregion

    /// <summary>
    /// Class implements the collection interface <see cref="IWorkItemCollection"/>.
    /// </summary>
    public class TfsWorkItemCollection : Collection<IWorkItem>, IWorkItemCollection
    {
        /// <summary>
        /// Retrieve a specific work item based on the work item id
        /// </summary>
        /// <param name="id">The work item id</param>
        /// <returns>The requested work item</returns>
        public IWorkItem Find(int id)
        {
            return Items.FirstOrDefault(x => x.Id == id);
        }
    }
}