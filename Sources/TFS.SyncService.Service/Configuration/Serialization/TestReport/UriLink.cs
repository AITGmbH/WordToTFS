using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// Configuration class for work item editor link.
    /// </summary>
    [XmlRoot("UriLink")]
    public class UriLink
    {
        /// <summary>
        /// Gets or sets the property to evaluate and gets the artifact uri of the work item editor.
        /// </summary>
        [XmlAttribute("Uri")]
        public string Uri { get; set; }
    }
}
