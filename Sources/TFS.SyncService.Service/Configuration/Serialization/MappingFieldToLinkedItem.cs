using System.Diagnostics.CodeAnalysis;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize one mapping 'field to linked item'.
    /// </summary>
    public class MappingFieldToLinkedItem
    {
        /// <summary>
        /// Property used to serialize 'LinkedWorkItemType' xml attribute.
        /// </summary>
        [XmlAttribute("LinkedWorkItemType")]
        public string LinkedWorkItemType { get; set; }

        /// <summary>
        /// Property used to serialize 'LinkType' xml attribute.
        /// </summary>
        /// <value>
        /// <c>Parent</c> - the linked work item is parent to the work item where is this configuration defined.
        /// <c>Child</c> - the linked work item is child to the work item where is this configuration defined.
        /// </value>
        [XmlAttribute("LinkType")]
        public LinkedItemLinkType LinkType { get; set; }

        /// <summary>
        /// Property used to serialize 'WorkItemBindType' xml attribute.
        /// </summary>
        [XmlAttribute("WorkItemBindType")]
        public WorkItemBindType WorkItemBindType { get; set; }

        /// <summary>
        /// Property used to serialize 'MappingFieldAssignment' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlArray("FieldAssignments")]
        [XmlArrayItem("FieldAssignment")]
        public MappingFieldAssignment[] FieldAssignments { get; set; }

        /// <summary>
        /// Property used to serialize 'MappingTableRow' xml attribute.
        /// </summary>
        /// <value>Value defines the row of the table,
        /// where is the source information stored to create the linked items.</value>
        [XmlAttribute("MappingTableRow")]
        public int TableRow { get; set; }

        /// <summary>
        /// Property used to serialize 'MappingTableCol' xml attribute.
        /// </summary>
        /// <value>Value defines the column of the table,
        /// where is the source information stored to create the linked items.</value>
        [XmlAttribute("MappingTableCol")]
        public int TableCol { get; set; }
    }
}
