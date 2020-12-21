using System;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;

namespace AIT.TFS.SyncService.Contracts.Adapter
{
    using System.Collections.Generic;
    using WorkItemObjects;

    /// <summary>
    /// Interface for all word adapter implementations. It extends the
    /// generic <c>IWorkItemSyncAdapter</c> with word specific methods
    /// </summary>
    public interface IWordSyncAdapter : IWorkItemSyncAdapter, IPreparationDocument
    {
        /// <summary>
        /// Gets the work items represented by the structures in the current selection
        /// </summary>
        /// <returns>The currently selected work items</returns>
        IEnumerable<IWorkItem> GetSelectedWorkItems();

        /// <summary>
        /// Gets the header that is currently selected
        /// </summary>
        /// <returns></returns>
        IEnumerable<IWorkItem> GetSelectedHeader(); 

        /// <summary>
        /// Creates a work item hyperlink at the current position in the document.
        /// </summary>
        /// <param name="text">Text for the hyperlink</param>
        /// <param name="uri">URI to the team server web access.</param>
        void CreateWorkItemHyperlink(string text, Uri uri);
    }
}
