namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for hyperlink - work item viewer.
    /// </summary>
    public interface IConfigurationTestWorkItemViewerLink
    {
        /// <summary>
        /// Gets or sets the property to evaluate and gets the id of the work item.
        /// </summary>
        string Id { get; }

        /// <summary>
        /// Gets or sets the property to evaluate and gets the revision of the work item.
        /// </summary>
        string Revision { get; }

        /// <summary>
        /// Gets the flag if the configured text for Hyperlink will be suppressed and text 'id - title' will be generated.
        /// </summary>
        bool AutoText { get; }

        /// <summary>
        /// Gets the string replacement formating of the <see cref="IConfigurationTestWorkItemViewerLink"/>
        /// </summary>
        string Format { get; }
    }
}
