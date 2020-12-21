using System.Diagnostics.CodeAnalysis;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize a header configuration.
    /// </summary>
    public class MappingHeader
    {
        /// <summary>
        /// Property used to serialize 'RelatedTemplate' xml attribute.
        /// </summary>
        [XmlAttribute("RelatedTemplate")]
        public string RelatedTemplate { get; set; }

        /// <summary>
        /// Property used to serialize 'Identifier' xml attribute.
        /// </summary>
        [XmlAttribute("Identifier")]
        public string Identifier { get; set; }

        /// <summary>
        /// Property used to serialize 'AssignTo' xml attribute.
        /// </summary>
        [XmlAttribute("AssignTo")]
        public string AssignTo { get; set; }

        /// <summary>
        /// Property used to serialize 'Row' xml attribute.
        /// </summary>
        [XmlAttribute("Row")]
        public int Row { get; set; }

        /// <summary>
        /// Property used to serialize 'Column' xml attribute.
        /// </summary>
        [XmlAttribute("Column")]
        public int Column { get; set; }

        /// <summary>
        /// Property used to serialize 'Level' xml attribute.
        /// </summary>
        [XmlAttribute("Level")]
        public int Level { get; set; }

        /// <summary>
        /// Property used to serialize 'Fields/Field' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlArray("Fields")]
        [XmlArrayItem("Field")]
        public MappingField[] Fields { get; set; }

        /// <summary>
        /// Property used to serialize 'Converters/Converter' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlArray("Converters")]
        [XmlArrayItem("Converter")]
        public MappingConverter[] Converters { get; set; }

        /// <summary>
        /// Property used to serialize image file name.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlAttribute("ImageFile")]
        public string ImageFile { get; set; }
    }
}