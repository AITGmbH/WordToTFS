using System.Diagnostics.CodeAnalysis;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize list of mapping settings.
    /// </summary>
    public class MappingList
    {
        /// <summary>
        /// Property used to serialize 'ExcludeRegularExpression' xml attribute.
        /// </summary>
        [XmlAttribute("ExcludeRegularExpression")]
        public string ExcludeRegularExpression { get; set; }

        /// <summary>
        /// Property used to serialize 'Mapping' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlElement("Mapping")]
        public MappingElement[] Mappings { get; set; }
    }
}