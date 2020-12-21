namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects
{
    using System;
    using Contracts.WorkItemObjects;

    /// <summary>
    /// Class is a holder of one error that occurred in save process.
    /// </summary>
    internal class SaveError : ISaveError
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SaveError"/> class.
        /// </summary>
        /// <param name="workItem">Work item where error on save occurred.</param>
        /// <param name="exception">Exception that occurred.</param>
        public SaveError(IWorkItem workItem, Exception exception)
        {
            WorkItem = workItem;
            Exception = exception;
        }

        /// <summary>
        /// Gets the associated <see cref="IWorkItem"/> with this error container.
        /// </summary>
        public IWorkItem WorkItem { get; private set; }

        /// <summary>
        /// Gets or sets associated <see cref="IField"/> with this error container.
        /// </summary>
        public IField Field { get; set; }

        /// <summary>
        /// Gets the occurred exception.
        /// </summary>
        public Exception Exception { get; private set; }
    }
}