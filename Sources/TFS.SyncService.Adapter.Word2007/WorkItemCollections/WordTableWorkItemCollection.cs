namespace AIT.TFS.SyncService.Adapter.Word2007.WorkItemCollections
{
    using System.Collections.ObjectModel;
    using System.Linq;
    using Contracts.WorkItemCollections;
    using Contracts.WorkItemObjects;

    /// <summary>
    /// Class implements the collection interface <see cref="IWorkItemCollection"/>.
    /// </summary>
    public class WordTableWorkItemCollection : Collection<IWorkItem>, IWorkItemCollection
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