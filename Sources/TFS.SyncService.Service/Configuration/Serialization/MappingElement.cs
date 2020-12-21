using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize one complete mapping object with field and conversion mappings.
    /// </summary>
    public class MappingElement
    {
        /// <summary>
        /// Related schema. The files name is assigned during reading operation.
        /// </summary>
        public string RelatedSchema { get; set; }

        /// <summary>
        /// Property used to serialize 'RelatedTemplate' xml attribute.
        /// </summary>
        [XmlAttribute("RelatedTemplate")]
        public string RelatedTemplate { get; set; }

        /// <summary>
        /// Property used to serialize 'WorkItemType' xml attribute.
        /// </summary>
        [XmlAttribute("WorkItemType")]
        public string WorkItemType { get; set; }

        /// <summary>
        /// Property used to serialize 'WorkItemSubtypeField' xml attribute.
        /// </summary>
        [XmlAttribute("WorkItemSubtypeField")]
        public string WorkItemSubtypeField { get; set; }

        /// <summary>
        /// Property used to serialize 'WorkItemSubtypeValue' xml attribute.
        /// </summary>
        [XmlAttribute("WorkItemSubtypeValue")]
        public string WorkItemSubtypeValue { get; set; }

        /// <summary>
        /// Property used to serialize 'MappingWorkItemType' xml attribute.
        /// </summary>
        [XmlAttribute("MappingWorkItemType")]
        public string MappingWorkItemType { get; set; }

        /// <summary>
        /// Property used to serialize 'AssignRegularExpression' xml attribute.
        /// </summary>
        [XmlAttribute("AssignRegularExpression")]
        public string AssignRegularExpression { get; set; }

        /// <summary>
        /// Property used to serialize 'AssignCellRow' xml attribute.
        /// </summary>
        [XmlAttribute("AssignCellRow")]
        public int AssignCellRow { get; set; }

        /// <summary>
        /// Property used to serialize 'AssignCellCol' xml attribute.
        /// </summary>
        [XmlAttribute("AssignCellCol")]
        public int AssignCellCol { get; set; }

        /// <summary>
        /// Property serialize "HideElementInWord" xml node.
        /// </summary>
        [XmlAttribute("HideElementInWord")]
        [DefaultValue(false)]
        public bool HideElementInWord { get; set; }

        /// <summary>
        /// Property used to serialize 'Fields/Field' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlArray("Fields")]
        [XmlArrayItem("Field")]
        public MappingField[] Fields { get; set; }

        /// <summary>
        /// Property used to serialize 'FieldsToLinkedItems/FieldToLinkedItem' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlArray("FieldsToLinkedItems")]
        [XmlArrayItem("FieldToLinkedItem")]
        public MappingFieldToLinkedItem[] FieldsToLinkedItems { get; set; }

        /// <summary>
        /// Property used to serialize 'Links/Link' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlArray("Links")]
        [XmlArrayItem("Link")]
        public MappingLink[] Links { get; set; }

        /// <summary>
        /// Property used to serialize 'Converters/Converter' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlArray("Converters")]
        [XmlArrayItem("Converter")]
        public MappingConverter[] Converters { get; set; }

        /// <summary>
        /// Property used to serialize image file name.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlAttribute("ImageFile")]
        public string ImageFile { get; set; }

        /// <summary>
        /// Property used to serialize the state transitions
        /// </summary>
        [XmlElement("Transitions")]
        public MappingStateTransition StateTransitions { get; set; }

        /// <summary>
        /// Property used to serialize pre-operations.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
        [XmlArray("PreOperations")]
        [XmlArrayItem("Operation")]
        public List<OperationConfiguration> PreOperations { get; set; }

        /// <summary>
        /// Property used to serialize post-operations
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
        [XmlArray("PostOperations")]
        [XmlArrayItem("Operation")]
        public List<OperationConfiguration> PostOperations { get; set; }
    }
}