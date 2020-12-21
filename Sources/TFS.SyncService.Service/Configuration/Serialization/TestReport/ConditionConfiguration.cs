using System.Collections.Generic;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// The class serializes configuration for one condition in template.
    /// </summary>
    [XmlRoot("Condition")]
    public class ConditionConfiguration
    {
        /// <summary>
        /// Gets or sets the property to evaluate and gets the value for condition.
        /// </summary>
        [XmlAttribute("Property")]
        public string PropertyToEvaluate { get; set; }

        /// <summary>
        /// Property used to serialize 'Values/Value' xml node.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
        [XmlArray("Values")]
        [XmlArrayItem("Value")]
        public List<string> Values { get; set; }
    }
}
