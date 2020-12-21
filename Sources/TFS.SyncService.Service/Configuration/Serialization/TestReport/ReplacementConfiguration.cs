using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using System.ComponentModel;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
	using AIT.TFS.SyncService.Contracts.Enums;

	/// <summary>
    /// The class serializes configuration for one replacement in template.
    /// </summary>
    [XmlRoot("Replacement")]
    public class ReplacementConfiguration
    {
        /// <summary>
        /// Gets or sets the name of bookmark to replace.
        /// </summary>
        [XmlAttribute("Bookmark")]
        public string Bookmark { get; set; }

        /// <summary>
        /// Gets or sets the property to evaluate and gets the value for replacement - used as text to show.
        /// </summary>
        [XmlAttribute("Property")]
        public string PropertyToEvaluate { get; set; }

        /// <summary>
        /// Gets or sets Parameters for the property to evaluate 
        /// </summary>
        [XmlAttribute("Parameters")]
        public string Parameters { get; set; }

		/// <summary>
		/// Gets the variable name if replacement should be based on a variable.
		/// </summary>
		[XmlAttribute("VariableName")]
		public string VariableName { get; set; }

		/// <summary>
		/// Property used to serialize 'FieldValueType' xml attribute.
		/// </summary>
		[XmlAttribute("FieldValueType")]
		[DefaultValue(FieldValueType.BasedOnVariable)]
		public FieldValueType FieldValueType { get; set; }

		/// <summary>
		/// Gets the type of value provided by evaluated of property.
		/// </summary>
		[XmlAttribute("ValueType")]
        [DefaultValue(PropertyValueFormat.PlainText)]
        public PropertyValueFormat ValueType { get; set; }

        /// <summary>
        /// Gets or sets the linked template to use for enumerable properties by replacement.
        /// </summary>
        [XmlAttribute("LinkedTemplate")]
        public string LinkedTemplate { get; set; }

        /// <summary>
        /// Gets or sets the configuration for work item editor link.
        /// </summary>
        [XmlElement("WorkItemEditorLink")]
        public WorkItemEditorLink WorkItemEditorLink { get; set; }

        /// <summary>
        /// Gets or sets the configuration for work item viewer link.
        /// </summary>
        [XmlElement("WorkItemViewerLink")]
        public WorkItemViewerLink WorkItemViewerLink { get; set; }

        /// <summary>
        /// Gets or sets the configuration for build viewer link.
        /// </summary>
        [XmlElement("BuildViewerLink")]
        public BuildViewerLink BuildViewerLink { get; set; }

        /// <summary>
        /// Gets or sets the configuration for common link.
        /// </summary>
        [XmlElement("UriLink")]
        public UriLink UriLink { get; set; }

        /// <summary>
        /// Gets or sets the configuration for attachment link.
        /// </summary>
        [XmlElement("AttachmentLink")]
        public AttachmentLink AttachmentLink { get; set; }

        [XmlAttribute("WordBookmark")]
        public string WordBookmark
        {
            get;
            set;
        }

        [XmlAttribute("ResolveResolutionState")]
        public bool ResolveResolutionState { get; set; }
    }
}
