using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// The interface defines the functionality of configuration for field to define linked work item.
    /// </summary>
    public interface IConfigurationFieldToLinkedItem
    {
        /// <summary>
        /// The property gets the name of work item that should be used as linked work item.
        /// </summary>
        string LinkedWorkItemType { get; }

        /// <summary>
        /// The property gets the type of the relationship between original work item and linked work item.
        /// </summary>
        /// <value>
        /// <c>Parent</c> - the linked work item is parent to the work item where is this configuration defined.
        /// <c>Child</c> - the linked work item is child to the work item where is this configuration defined.
        /// </value>
        LinkedItemLinkType LinkType { get; }

        /// <summary>
        /// The property gets the kind of definition for the linked work item.
        /// </summary>
        WorkItemBindType WorkItemBindType { get; }

        /// <summary>
        /// The property gets all configured fields of type <see cref="IConfigurationFieldToLinkedItem"/> to define linked work items.
        /// </summary>
        IList<IConfigurationFieldAssignment> FieldAssignmentConfiguration { get; }

        /// <summary>
        /// The property gets the index of column in the table where is the definition for the linked work items.
        /// </summary>
        int ColIndex { get; }

        /// <summary>
        /// The property gets the index of row in the table where is the definition for the linked work items.
        /// </summary>
        int RowIndex { get; }
    }
}
