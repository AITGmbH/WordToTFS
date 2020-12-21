using System.Diagnostics.CodeAnalysis;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize one mapping converter node.
    /// </summary>
    public class MappingConverter
    {
        /// <summary>
        /// Property used to serialize 'FieldName' xml attribute.
        /// </summary>
        [XmlAttribute("FieldName")]
        public string FieldName { get; set; }

        /// <summary>
        /// Property used to serialize 'Values/Value' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlArrayItem("Value")]
        public MappingConverterValue[] Values { get; set; }
    }
}