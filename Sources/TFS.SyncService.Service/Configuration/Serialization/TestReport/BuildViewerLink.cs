using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// Configuration class for build viewer link.
    /// </summary>
    [XmlRoot("BuildViewerLink")]
    public class BuildViewerLink
    {
        /// <summary>
        /// Gets or sets the property to evaluate and gets the build number of the build.
        /// </summary>
        [XmlAttribute("BuildNumber")]
        public string BuildNumber { get; set; }
    }
}
