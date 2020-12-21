using System.ComponentModel;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// Configuration class for work item editor link.
    /// </summary>
    [XmlRoot("AttachmentLink")]
    public class AttachmentLink
    {
        /// <summary>
        /// Gets or sets how attachment links should be treated
        /// </summary>
        [DefaultValue(AttachmentLinkMode.DownloadAndLinkToLocalFile)]
        [XmlAttribute("Mode")]
        public AttachmentLinkMode Mode { get; set; }
    }
}