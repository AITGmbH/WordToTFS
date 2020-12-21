using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize one mapping field.
    /// </summary>
    public class MappingLink
    {
        /// <summary>
        /// Property used to serialize 'FieldValueType' xml attribute.
        /// </summary>
        [XmlAttribute("Type")]
        public string LinkValueType
        {
            get;
            set;
        }

        /// <summary>
        /// Property used to serialize 'Direction' xml attribute.
        /// </summary>
        [XmlAttribute("Direction")]
        public Direction Direction
        {
            get;
            set;
        }

        /// <summary>
        /// Property used to serialize 'MappingTableRow' xml attribute.
        /// </summary>
        [XmlAttribute("MappingTableRow")]
        public int TableRow
        {
            get;
            set;
        }

        /// <summary>
        /// Property used to serialize 'MappingTableCol' xml attribute.
        /// </summary>
        [XmlAttribute("MappingTableCol")]
        public int TableCol
        {
            get;
            set;
        }

        /// <summary>
        /// Property used to serialize 'Overwrite' xml attribute.
        /// </summary>
        [XmlAttribute("Overwrite")]
        public bool Overwrite
        {
            get;
            set;
        }

        /// <summary>
        /// Property used to serialize 'Overwrite' xml attribute.
        /// </summary>
        [XmlAttribute("LinkedWorkItemTypes")]
        public string LinkedWorkItemTypes
        {
            get;
            set;
        }

        /// <summary>
        /// Property used to serialize 'LinkSeparator' xml attribute
        /// </summary>
        [XmlAttribute("LinkSeparator")]
        public string LinkSeparator
        {
            get;
            set;
        }

        /// <summary>
        /// Property used to serialize 'LinkFormat' xml attribute
        /// </summary>
        [XmlAttribute("LinkFormat")]
        public string LinkFormat
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the type of work item to which to link automatically when publishing
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "LinkWork")]
        [XmlAttribute("AutoLinkWorkItemType")]
        public string AutomaticLinkWorkItemType
        {
            get;
            set;
        }

        /// <summary>
        /// Gets an optional field used to select the automatically linked work item
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "LinkWork")]
        [XmlAttribute("AutoLinkWorkItemSubtypeField")]
        public string AutomaticLinkWorkItemSubtypeField
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the value the optional field has to have
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "LinkWork")]
        [XmlAttribute("AutoLinkWorkItemSubtypeValue")]
        public string AutomaticLinkWorkItemSubtypeValue
        {
            get;
            set;
        }

        /// <summary>
        /// Gets a value indicating whether to suppress warnings when no automatic link target was found
        /// </summary>
        [XmlAttribute("AutoLinkSuppressWarnings")]
        public bool AutomaticLinkSuppressWarnings
        {
            get;
            set;
        }
    }
}