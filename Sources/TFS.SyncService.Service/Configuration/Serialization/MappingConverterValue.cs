using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize one value for mapping converter.
    /// </summary>
    public class MappingConverterValue
    {
        /// <summary>
        /// Property used to serialize 'Text' xml attribute.
        /// </summary>
        [XmlAttribute("Text")]
        public string Text { get; set; }

        /// <summary>
        /// Property used to serialize 'MappingText' xml attribute.
        /// </summary>
        [XmlAttribute("MappingText")]
        public string MappingText { get; set; }
    }
}