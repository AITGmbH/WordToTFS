namespace AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects
{
    using System;
    using Contracts.WorkItemObjects;

    /// <summary>
    /// Class is a holder of one error that occurred in save process.
    /// </summary>
    internal class SaveError : ISaveError
    {
        /// <summary>
        /// Gets the associated <see cref="IWorkItem"/> with this error container.
        /// </summary>
        public IWorkItem WorkItem { get; set; }

        /// <summary>
        /// Gets or sets associated <see cref="IField"/> with this error container.
        /// </summary>
        public IField Field { get; set; }

        /// <summary>
        /// Gets the occurred exception.
        /// </summary>
        public Exception Exception { get; set; }
    }
}
