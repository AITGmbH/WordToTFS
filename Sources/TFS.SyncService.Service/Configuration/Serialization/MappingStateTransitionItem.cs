namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    using System.Xml.Serialization;

    /// <summary>
    /// A Single entry in a state transition table
    /// </summary>
    public class MappingStateTransitionItem
    {
        /// <summary>
        /// Source state of this transition
        /// </summary>
        [XmlAttribute("From")]
        public string From { get; set; }

        /// <summary>
        /// Target state of this transition
        /// </summary>
        [XmlAttribute("To")]
        public string To { get; set; }
    }
}
