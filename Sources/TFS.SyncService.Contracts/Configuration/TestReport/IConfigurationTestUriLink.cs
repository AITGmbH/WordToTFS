namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for common hyperlink.
    /// </summary>
    public interface IConfigurationTestUriLink
    {
        /// <summary>
        /// Gets or sets the property to evaluate and gets the common uri.
        /// </summary>
        string Uri { get; }
    }
}
