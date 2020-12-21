using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects
{
    /// <summary>
    /// Wrapper for supported link types.
    /// </summary>
    public class TfsWorkItemLinkType : ITFSWorkItemLinkType
    {
        private readonly WorkItemLinkType _workItemLinkType;
        private readonly LinkType _linkType;

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsWorkItemLinkType"/> class.
        /// </summary>
        /// <param name="workItemLinkType">Type of the work item link.</param>
        /// <param name="linkType">Type of the link.</param>
        internal TfsWorkItemLinkType(WorkItemLinkType workItemLinkType, LinkType linkType)
        {
            _workItemLinkType = workItemLinkType;
            _linkType = linkType;
        }

        /// <summary>
        /// Gets the type of the link.
        /// </summary>
        public LinkType LinkType
        {
            get { return _linkType; }
        }

        /// <summary>
        /// Gets the underlying <see cref="WorkItemLinkType" />.
        /// </summary>
        public WorkItemLinkType WorkItemLinkType
        {
            get { return _workItemLinkType; }
        }

        /// <summary>
        /// Gets a user friendly name for the link type.
        /// </summary>
        public string DisplayName
        {
            get
            {
                switch(_linkType)
                {
                    case LinkType.Full:
                        return _workItemLinkType.ForwardEnd.Name + "/" + _workItemLinkType.ReverseEnd.Name;

                    case LinkType.ForwardEnd:
                        return _workItemLinkType.ForwardEnd.Name;

                    case LinkType.ReverseEnd:
                        return _workItemLinkType.ReverseEnd.Name;

                    default:
                        return string.Empty;
                }
            }
        }

        /// <summary>
        /// Gets the internal reference name of the link type.
        /// </summary>
        public string ReferenceName
        {
            get
            {
                string referenceName = string.Empty;
                switch (_linkType)
                {
                    case LinkType.Full:
                        referenceName = _workItemLinkType.ReferenceName;
                        break;
                    case LinkType.ForwardEnd:
                        referenceName = _workItemLinkType.ForwardEnd.Name;
                        break;
                    case LinkType.ReverseEnd:
                        referenceName = _workItemLinkType.ReverseEnd.Name;
                        break;
                }

                return referenceName;
            }
        }

        /// <summary>
        /// Gets a wrapper for the given work item link type.
        /// </summary>
        /// <param name="store">The work item store in which to search the link type.</param>
        /// <param name="referenceName">Reference name of the link type.</param>
        /// <returns>Wrapper for the found link type or null if no link type with that name exists.</returns>
        /// <exception cref="System.ArgumentNullException">store</exception>
        public static ITFSWorkItemLinkType GetTfsWorkItemLinkTypeByReferenceName(WorkItemStore store, string referenceName)
        {
            Guard.ThrowOnArgumentNull(store, "store");

            foreach (var workItemLinkType in store.WorkItemLinkTypes)
            {
                if (workItemLinkType.ReferenceName == referenceName)
                {
                    return new TfsWorkItemLinkType(workItemLinkType, LinkType.Full);
                }
                
                if (workItemLinkType.ForwardEnd.Name == referenceName)
                {
                    return new TfsWorkItemLinkType(workItemLinkType, LinkType.ForwardEnd);
                }

                if (workItemLinkType.ReverseEnd.Name == referenceName)
                {
                    return new TfsWorkItemLinkType(workItemLinkType, LinkType.ReverseEnd);
                }
            }

            return null;
        }
    }
}
