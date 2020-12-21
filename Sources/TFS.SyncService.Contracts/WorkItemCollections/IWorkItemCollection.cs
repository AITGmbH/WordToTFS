namespace AIT.TFS.SyncService.Contracts.WorkItemCollections
{
    using System.Collections.Generic;
    using WorkItemObjects;

    /// <summary>
    /// Simple extension of generic collection
    /// </summary>
    public interface IWorkItemCollection : ICollection<IWorkItem>
    {
        /// <summary>
        /// Retrieve a specific work item based on the work item id
        /// </summary>
        /// <param name="id">The work item id</param>
        /// <returns>The requested work item</returns>
        IWorkItem Find(int id);
    }
}