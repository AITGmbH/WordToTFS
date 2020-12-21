using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize one field in 'field to linked item'.
    /// </summary>
    public class MappingFieldAssignment
    {
        /// <summary>
        /// Property used to serialize 'Name' xml attribute.
        /// </summary>
        [XmlAttribute("Name")]
        public string Name { get; set; }

        /// <summary>
        /// Property used to serialize 'MappingName' xml attribute.
        /// </summary>
        [XmlAttribute("MappingName")]
        public string MappingName { get; set; }

        /// <summary>
        /// Property used to serialize 'FieldPosition' xml attribute.
        /// </summary>
        [XmlAttribute("FieldPosition")]
        public FieldPositionType FieldPosition { get; set; }

        /// <summary>
        /// Property used to serialize 'FieldValueType' xml attribute.
        /// </summary>
        [XmlAttribute("FieldValueType")]
        public FieldValueType FieldValueType { get; set; }

        /// <summary>
        /// Property used to serialize 'Direction' xml attribute.
        /// </summary>
        [XmlAttribute("Direction")]
        public Direction Direction { get; set; }
    }
}
