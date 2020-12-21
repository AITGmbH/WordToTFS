using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// QueryConfiguration that is passes to the TFS adapter to define what work items to load
    /// </summary>
    public class QueryConfiguration : IQueryConfiguration
    {
        /// <summary>
        /// Create new instance with no special configuration.
        /// </summary>
        public QueryConfiguration()
        {
            LinkTypes = new List<string>();
            ByIDs = new List<int>();
        }

        /// <summary>
        /// Path to a stored query. Used if <see cref="ImportOption"/> is set to <see cref="QueryImportOption.SavedQuery"/>
        /// </summary>
        public string QueryPath { get; set; }

        /// <summary>
        /// Determines whether to also load linked work items. Ignored if <see cref="ImportOption"/> is set to <see cref="QueryImportOption.SavedQuery"/>
        /// </summary>
        public bool UseLinkedWorkItems { get; set; }

        /// <summary>
        /// Determines whether to get only direct links only or recursive
        /// </summary>
        public bool IsDirectLinkOnlyMode { get; set; }

        /// <summary>
        /// The link types to include when <see cref="UseLinkedWorkItems"/> is true
        /// </summary>
        public ICollection<string> LinkTypes { get; private set; }

        /// <summary>
        /// Defines by what means to load work items. Load work items using a saved query or identify them by id or title.
        /// </summary>
        public QueryImportOption ImportOption { get; set; }

        /// <summary>
        /// A list of ids of work items to load. Used only if <see cref="ImportOption"/> is set to <see cref="QueryImportOption.IDs"/>
        /// </summary>
        public ICollection<int> ByIDs { get; set; }

        /// <summary>
        /// A substring of the title of work items to retrieve. Used only if <see cref="ImportOption"/> is set to <see cref="QueryImportOption.TitleContains"/>
        /// </summary>
        public string ByTitle { get; set; }

        /// <summary>
        /// Set to a work item type to restrict search by title to that type. Use "" or null to search all work item types.
        /// Used only if <see cref="ImportOption"/> is set to <see cref="QueryImportOption.TitleContains"/>
        /// </summary>
        public string WorkItemType { get; set; }
    }
}