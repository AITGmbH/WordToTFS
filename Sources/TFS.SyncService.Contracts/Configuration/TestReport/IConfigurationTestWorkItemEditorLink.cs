namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for hyperlink - work item editor.
    /// </summary>
    public interface IConfigurationTestWorkItemEditorLink
    {
        /// <summary>
        /// Gets or sets the property to evaluate and gets the artifact uri of the work item editor.
        /// </summary>
        string Uri { get; }
    }
}
