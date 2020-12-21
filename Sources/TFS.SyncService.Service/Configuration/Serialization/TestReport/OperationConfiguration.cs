using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// Configuration class for one operation before / after insert of template.
    /// </summary>
    [XmlRoot("Operation")]
    public class OperationConfiguration
    {
        /// <summary>
        /// Gets or sets the type of operation to execute before / after insert of template.
        /// </summary>
        [XmlAttribute("Type")]
        public OperationType OperationType { get; set; }
    }
}
