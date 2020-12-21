using System.ComponentModel;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// Configuration class for work item viewer link.
    /// </summary>
    [XmlRoot("WorkItemViewerLink")]
    public class WorkItemViewerLink
    {
        /// <summary>
        /// Gets or sets the property to evaluate and gets the id of the work item.
        /// </summary>
        [XmlAttribute("Id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the property to evaluate and gets the revision of the work item.
        /// </summary>
        [XmlAttribute("Revision")]
        public string Revision { get; set; }

        /// <summary>
        /// Gets the flag if the configured text for Hyperlink will be suppressed and text 'id - title' will be generated.
        /// </summary>
        [XmlAttribute("AutoText")]
        [DefaultValue(false)]
        public bool AutoText { get; set; }

        [XmlAttribute("Format")]
        public string Format { get; set; }
    }
}
