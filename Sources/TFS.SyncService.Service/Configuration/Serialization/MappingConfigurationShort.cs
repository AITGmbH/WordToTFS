using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used as top serialization class to serialize only base mapping bundle information.
    /// </summary>
    [XmlRoot("MappingConfiguration", Namespace = "", IsNullable = false)]
    [XmlType("MappingConfiguration", Namespace = "")]
    public class MappingConfigurationShort
    {
        /// <summary>
        /// Property used to serialize 'ShowName' xml attribute.
        /// </summary>
        [XmlAttribute("ShowName")]
        public string ShowName { get; set; }

        /// <summary>
        /// Property used to serialize 'RelatedSchema' xml attribute.
        /// </summary>
        [XmlAttribute("RelatedSchema")]
        public string RelatedSchema { get; set; }

        /// <summary>
        /// Property used to serialize 'DefaultMapping' xml attribute.
        /// </summary>
        [XmlAttribute("DefaultMapping")]
        public bool DefaultMapping { get; set; }

        /// <summary>
        /// Property serialize "UseStackRank" xml node
        /// </summary>
        [XmlAttribute("UseStackRank")]
        public bool UseStackRank { get; set; }

        /// <summary>
        /// Property serialize "TypeOfHierarchyRelationships" xml node.
        /// </summary>
        [XmlAttribute("TypeOfHierarchyRelationships")]
        public string TypeOfHierarchyRelationships { get; set; }
    }
}