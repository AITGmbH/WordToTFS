using System.Diagnostics.CodeAnalysis;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    using System.Xml.Serialization;

    /// <summary>
    /// State transition table
    /// </summary>
    public class MappingStateTransition
    {
        /// <summary>
        /// Field to which these state transitions apply
        /// </summary>
        [XmlAttribute("FieldName")]
        public string FieldName { get; set; }

        /// <summary>
        /// An array of defined state transitions
        /// </summary>
        [XmlElement("Transition")]
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        public MappingStateTransitionItem[] Items { get; set; }
    }
}
