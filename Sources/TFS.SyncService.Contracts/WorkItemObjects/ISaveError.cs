using System;

namespace AIT.TFS.SyncService.Contracts.WorkItemObjects
{
    /// <summary>
    /// Interface defines the error container associated with <see cref="IWorkItem"/>.
    /// </summary>
    public interface ISaveError
    {
        /// <summary>
        /// Gets the associated <see cref="IWorkItem"/> with this error container.
        /// </summary>
        IWorkItem WorkItem { get; }

        /// <summary>
        /// Gets or sets associated <see cref="IField"/> with this error container.
        /// </summary>
        IField Field { get; set; }

        /// <summary>
        /// Gets the occurred exception.
        /// </summary>
        Exception Exception { get; }
    }
}