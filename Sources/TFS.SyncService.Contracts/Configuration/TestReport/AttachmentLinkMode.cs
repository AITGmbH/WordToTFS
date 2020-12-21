namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Allowed modes to define how attachments links should be handled
    /// </summary>
    public enum AttachmentLinkMode
    {
        /// <summary>
        /// Downloads the attachments and creates a link to the local file
        /// </summary>
        DownloadAndLinkToLocalFile,

        /// <summary>
        /// Downloads the attachments and creates a link to the local file, whereas the attachmentfolder doesn's contain the GUID
        /// </summary>
        DownloadAndLinkToLocalFileWithoutGuid,

        /// <summary>
        /// Links to the server version of the attachment
        /// </summary>
        LinkToServerVersion,

        /// <summary>
        /// Downloads the attachment but does not link to it
        /// </summary>
        DownloadOnly
    }
}
