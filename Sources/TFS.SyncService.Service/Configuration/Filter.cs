using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Service.Configuration
{

    /// <summary>
    /// Filter for for link types
    /// </summary>
    public class Filter : IFilter
    {
        /// <summary>
        /// specifies a link type for the filter
        /// </summary>
        public string LinkType
        {
            get;
            set;
        }

        /// <summary>
        /// Filter for work items
        /// </summary>
        public string FilterOn
        {
            get;
            set;
        }
    }
}
