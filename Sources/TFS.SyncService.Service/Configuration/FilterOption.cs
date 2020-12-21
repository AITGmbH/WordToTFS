using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Service.Configuration
{

    /// <summary>
    /// Filter optoions for the object queries
    /// </summary>
    public class FilterOption : IFilterOption
    {

        /// <summary>
        /// The distinct filter
        /// </summary>
        public bool Distinct
        {
            get;
            set;
        }

        /// <summary>
        /// Show only latest test results
        /// </summary>
        public bool Latest
        {
            get;
            set;
        }

        /// <summary>
        /// Filter for a property
        /// </summary>
        public string FilterProperty
        {
            get;
            set;
        }
    }
}
