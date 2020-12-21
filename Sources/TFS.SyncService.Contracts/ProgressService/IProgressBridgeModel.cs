namespace AIT.TFS.SyncService.Contracts.ProgressService
{
    /// <summary>
    /// Interface defines functionality of bridge to GUI.
    /// </summary>
    public interface IProgressBridgeModel
    {
        /// <summary>
        /// Gets or sets the title of the whole progress process.
        /// </summary>
        string ProgressTitle { get; set; }

        /// <summary>
        /// Gets or sets the value of progress. Value is from interval &lt;0, 100&gt;.
        /// </summary>
        int ProgressValue { get; set; }

        /// <summary>
        /// Gets or sets the text of progress.
        /// </summary>
        string ProgressText { get; set; }

        /// <summary>
        /// Gets the information if the long operation is to cancel.
        /// </summary>
        /// <returns>True for cancel the progress, otherwise false.</returns>
        bool ProgressCanceled { get; }
    }
}