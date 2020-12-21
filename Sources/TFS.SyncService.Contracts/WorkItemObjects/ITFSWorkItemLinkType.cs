using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace AIT.TFS.SyncService.Contracts.WorkItemObjects
{
    /// <summary>
    /// Possible link types of work item links. For example a child is the forward end of a parent-child relationship while the parent is the backward end.
    /// </summary>
    public enum LinkType
    {
        /// <summary>
        /// Link is non-directional.
        /// </summary>
        Full,

        /// <summary>
        /// Link is forward end of a directional link.
        /// </summary>
        ForwardEnd,

        /// <summary>
        /// Link is reverse end of a directional link.
        /// </summary>
        ReverseEnd,
    }

    /// <summary>
    /// Interface for a supported work item link type that can exist between two work items.
    /// </summary>
    public interface ITFSWorkItemLinkType
    {
        /// <summary>
        /// Gets the type of the link.
        /// </summary>
        LinkType LinkType { get; }

        /// <summary>
        /// Gets the underlying <see cref="WorkItemLinkType"/>.
        /// </summary>
        WorkItemLinkType WorkItemLinkType { get; }

        /// <summary>
        /// Gets a user friendly name for the link type.
        /// </summary>
        string DisplayName { get; }

        /// <summary>
        /// Gets the internal reference name of the link type.
        /// </summary>
        string ReferenceName { get; }
    }
}
