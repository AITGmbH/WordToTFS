using System.ComponentModel;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used to serialize one mapping field.
    /// </summary>
    public class MappingField
    {
        /// <summary>
        /// Property used to serialize 'Name' xml attribute.
        /// </summary>
        [XmlAttribute("Name")]
        public string Name { get; set; }

        /// <summary>
        /// Property used to serialize 'MappingName' xml attribute.
        /// </summary>
        [XmlAttribute("MappingName")]
        public string MappingName { get; set; }

        /// <summary>
        /// Property used to serialize 'FieldValueType' xml attribute.
        /// </summary>
        [XmlAttribute("FieldValueType")]
        [DefaultValue(FieldValueType.PlainText)]
        public FieldValueType FieldValueType { get; set; }

        /// <summary>
        /// Property used to serialize 'Format' xml attribute.
        /// </summary>
        [XmlAttribute("")]
        public string DateTimeFormat { get; set; }

        /// <summary>
        /// Property used to serialize 'Direction' xml attribute.
        /// </summary>
        [XmlAttribute("Direction")]
        public Direction Direction { get; set; }

        /// <summary>
        /// Property used to serialize 'MappingTableRow' xml attribute.
        /// </summary>
        [XmlAttribute("MappingTableRow")]
        public int TableRow { get; set; }

        /// <summary>
        /// Property used to serialize 'MappingTableCol' xml attribute.
        /// </summary>
        [XmlAttribute("MappingTableCol")]
        public int TableCol { get; set; }

        /// <summary>
        /// Property used to serialize "TestCaseStepDelimiter" xml attribute.
        /// </summary>
        [XmlAttribute("TestCaseStepDelimiter")]
        public string TestCaseStepDelimiter { get; set; }

        /// <summary>
        /// Property used to serialize "HandleAsDocument" xml attribute.
        /// </summary>
        [XmlAttribute("HandleAsDocument")]
        [DefaultValue(false)]
        public bool HandleAsDocument { get; set; }

        /// <summary>
        /// Property used to serialize "HandleAsDocumentMode" xml attribute.
        /// </summary>
        [XmlAttribute("HandleAsDocumentMode")]
        [DefaultValue(HandleAsDocumentType.OleOnDemand)]
        public HandleAsDocumentType HandleAsDocumentMode { get; set; }

        /// <summary>
        /// Property used to serialize "OLEMarkerField" xml attribute.
        /// </summary>
        [XmlAttribute("OLEMarkerField")]
        public string OLEMarkerField { get; set; }

        /// <summary>
        /// Property used to serialize "OLEMarkerValue" xml attribute.
        /// </summary>
        [XmlAttribute("OLEMarkerValue")]
        public string OLEMarkerValue { get; set; }

        /// <summary>
        /// Property used to serialize 'DefaultValue' xml node.
        /// </summary>
        [XmlElement("DefaultValue")]
        public MappingFieldDefaultValue DefaultValue { get; set; }

        /// <summary>
        /// Gets or sets the shape only workaround mode.
        /// </summary>
        [XmlAttribute("ShapeOnlyWorkaroundMode")]
        [DefaultValue(ShapeOnlyWorkaroundMode.AddSpace)]
        public ShapeOnlyWorkaroundMode ShapeOnlyWorkaroundMode
        {
            get;
            set;
        }
         [XmlAttribute("WordBookmark")]
        public string WordBookmark
        {
            get;
            set;
        }

         [XmlAttribute("VariableName")]
         public string VariableName
         {
             get;
             set;
         }
    }
}