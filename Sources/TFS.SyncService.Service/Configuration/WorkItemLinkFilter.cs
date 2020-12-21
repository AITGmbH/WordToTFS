using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Service.Configuration
{

    /// <summary>
    /// Represents the workitemfilter of the configuration class
    /// </summary>
    public class WorkItemLinkFilter : IWorkItemLinkFilter
    {

        /// <summary>
        /// Public constructor
        /// </summary>
        public WorkItemLinkFilter ()
        {
            Filters = new List<IFilter>();
        }

        /// <summary>
        /// The different filter types
        /// </summary>
        public FilterType FilterType
        {
            get;
            set;
        }

        /// <summary>
        /// The list of filters
        /// </summary>
        public IList<IFilter> Filters
        {
            get;
            set;
        }
    }
}
