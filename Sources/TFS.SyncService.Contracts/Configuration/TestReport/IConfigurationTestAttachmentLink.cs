namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Defines the properties of an attachment link in a generated document
    /// </summary>
    public interface IConfigurationTestAttachmentLink
    {
        /// <summary>
        /// Gets or sets how attachment links should be treated
        /// </summary>
        AttachmentLinkMode Mode { get; }
    }
}