using AIT.TFS.SyncService.Contracts.WorkItemCollections;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Contracts.WorkItemObjects
{
    /// <summary>
    /// Interface defines the functionality of a work item.
    /// </summary>
    public interface IWorkItem
    {
        /// <summary>
        /// Gets the work item id of the current work item. This value has to be unique among all work items
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets the work item revision.
        /// </summary>
        int Revision { get; }
            
        /// <summary>
        /// Gets work item title.
        /// </summary>
        string Title { get;  }

        /// <summary>
        /// Gets the type of the work item
        /// </summary>
        string WorkItemType { get; }

        /// <summary>
        /// Gets all fields which are supported by the current work item
        /// </summary>
        IFieldCollection Fields { get; }

        /// <summary>
        /// Ids of linked work items, grouped by link type
        /// </summary>
        Dictionary<IConfigurationLinkItem, int[]> Links { get; }

        /// <summary>
        /// Returns whether this work item has been modified
        /// </summary>
        bool IsDirty { get; }

        /// <summary>
        /// Returns whether this work item is new
        /// </summary>
        bool IsNew { get; }

        /// <summary>
        /// Creates a bookmark for this work item.
        /// </summary>
        void CreateBookmark();

        /// <summary>
        /// Return wheter this work items had errors during the validation on the word side
        /// </summary>
        bool HasWordValidationError
        {
            get;
            set;
        }

        /// <summary>
        /// Adds multiple links to other work items to this work item.
        /// </summary>
        /// <param name="adapter">Adapter used to look up more information about the referenced work items.</param>
        /// <param name="ids">The ids of the other work items.</param>
        /// <param name="linkType">The reference name of the link type.</param>
        /// <param name="overwrite">Sets whether the passed links replace all existing links of the given type.</param>
        /// <returns>True when the link was successfully added, false otherwise.</returns>
        bool AddLinks(IWorkItemSyncAdapter adapter, int[] ids, string linkType, bool overwrite);

        /// <summary>
        /// Adds multiple links to other work items to this work item.
        /// </summary>
        /// <param name="adapter">Adapter used to look up more information about the referenced work items.</param>
        /// <param name="ids">The ids of the other work items.</param>
        /// <param name="linkType">The reference name of the link type.</param>
        /// <param name="linkedWorkItemTypes">Determines a comma separated list of Work Item Types that can be used to filter the links in addition to the linkValueType.</param>
        /// <param name="overwrite">Sets whether the passed links replace all existing links of the given type.</param>
        /// <returns>True when the link was successfully added, false otherwise.</returns>
        bool AddLinks(IWorkItemSyncAdapter adapter, int[] ids, string linkType, string linkedWorkItemTypes, bool overwrite);

        /// <summary>
        /// Refresh work item to get latest state.
        /// </summary>
        void Refresh();
        
        /// <summary>
        /// Get revision number for this work item. If the referenceFieldName is defined then:
        /// TFS: Gets revision number which belongs to the field -> last revision number for certain field.
        /// Word: Gets revision number as value from reference field name, if parameter not defined then in default use System.Rev.
        /// </summary>
        /// <param name="referenceFieldName">Reference field name.</param>
        /// <returns>Work item revision with the most recent change to the field.</returns>
        int GetFieldRevision(string referenceFieldName);

        /// <summary>
        /// Gets a list of fields and their values for a given revision of the work item.
        /// </summary>
        /// <param name="revision">Revision for which to retrieve values.</param>
        /// <returns>List of fields for this revision.</returns>
        IFieldCollection GetWorkItemByRevision(int revision);

        /// <summary>
        /// Gets the configuration of this work item and all its fields
        /// </summary>
        IConfigurationItem Configuration { get; }

        /// <summary>
        /// Returns whether this and another work item are wrapper for the same work item.
        /// </summary>
        /// <param name="other">Other work item.</param>
        /// <returns>True, if both work item wrappers are wrapper for the same work item.</returns>
        bool IsSameWorkItem(IWorkItem other);

        /// <summary>
        /// Delete Id and Revision
        /// </summary>
        void DeleteIdAndRevision();
       
    }
}