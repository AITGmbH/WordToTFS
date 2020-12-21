using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize one default value of field.
    /// </summary>
    public class MappingFieldDefaultValue
    {
        /// <summary>
        /// Property used to serialize 'ShowName' xml attribute.
        /// </summary>
        [XmlAttribute("ShowName")]
        public string ShowName { get; set; }

        /// <summary>
        /// Property used to serialize the node value.
        /// </summary>
        [XmlText]
        public string DefaultValue { get; set; }
    }
}