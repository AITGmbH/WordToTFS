using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.WorkItemCollections;

namespace AIT.TFS.SyncService.Contracts.WorkItemObjects
{
    /// <summary>
    /// The interface defines the functionality of the object with support of linked work items.
    /// </summary>
    public interface IWorkItemLinkedItems : IWorkItem
    {
        /// <summary>
        /// The property gets the Dictionary of all linked items ordered in 'link type' groups.
        /// Key of dictionary means the link type and the values are all work to link with mentioned link type.
        /// </summary>
        IDictionary<string, IWorkItemCollection> LinkedWorkItems { get; }
    }
}
