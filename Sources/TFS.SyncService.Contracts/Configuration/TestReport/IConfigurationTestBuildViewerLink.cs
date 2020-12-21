namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for hyperlink - build viewer.
    /// </summary>
    public interface IConfigurationTestBuildViewerLink
    {
        /// <summary>
        /// Gets or sets the property to evaluate and gets the build number of the build.
        /// </summary>
        string BuildNumber { get; }
    }
}
