using System.ComponentModel;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize one mapping field.
    /// </summary>
    public class ButtonCustomization : IConfigurationButtonCustomization
    {
        /// <summary>
        /// Property used to serialize 'Name' xml attribute.
        /// </summary>
        [XmlAttribute("Name")]
        public string Name
        {
            get;
            set;
        }

        /// <summary>
        /// Property used to serialize 'MappingName' xml attribute.
        /// </summary>
        [XmlAttribute("Text")]
        public string Text
        {
            get;
            set;
        }
    }
}