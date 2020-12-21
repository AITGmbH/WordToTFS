using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationTestAttachmentLink"/>.
    /// </summary>
    internal class ConfigurationAttachmentLink : IConfigurationTestAttachmentLink
    {
        /// <summary>
        /// Gets or sets how attachment links should be treated
        /// </summary>
        public AttachmentLinkMode Mode { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationAttachmentLink"/> class.
        /// </summary>
        /// <param name="attachmentLink">The attachment link.</param>
        public ConfigurationAttachmentLink(AttachmentLink attachmentLink)
        {
            Mode = attachmentLink.Mode;
        }
    }
}