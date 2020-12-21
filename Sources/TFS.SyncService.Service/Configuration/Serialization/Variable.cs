using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    public class Variable : IVariable
    {
        /// <summary>
        /// Property used to serialize 'name' xml attribute.
        /// </summary>
        [XmlAttribute("Name")]
        public string Name { get; set; }


        /// <summary>
        /// Property used to serialize 'name' xml attribute.
        /// </summary>
        [XmlAttribute("Value")]
        public string Value { get; set; }
    }
}
