using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    ///Base class of build filters.
    /// </summary>
    [XmlRoot("BuildFilters")]
    public class BuildFilters : IBuildFilters
    {
        /// <summary>
        /// Gets the list of configured build qualities to use as filter for builds.
        /// </summary>
        [XmlArray("BuildQualities")]
        [XmlArrayItem("BuildQuality")]
        public List<string> BuildQualities { get; set; }

        /// <summary>
        /// Gets the list of configured build names to use as filter for builds.
        /// </summary>
        [XmlArray("BuildNames")]
        [XmlArrayItem("BuildName")]
        public List<string> BuildNames { get; set; }

        /// <summary>
        /// Gets the list of configured build ages to use as filter for builds.
        /// </summary>
        [XmlElement("BuildAge")]
        public string BuildAge { get; set; }

        /// <summary>
        /// Gets the list of configured build tags to use as filter for builds.
        /// </summary>
        [XmlArray("BuildTags")]
        [XmlArrayItem("BuildTag")]
        public List<string> BuildTags { get; set; }



    }
}
