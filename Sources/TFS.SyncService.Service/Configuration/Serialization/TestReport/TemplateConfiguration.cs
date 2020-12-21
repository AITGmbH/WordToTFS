using System.Collections.Generic;
using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
{
    /// <summary>
    /// <see cref="TemplateConfiguration"/> is configuration class for all configured templates.
    /// </summary>
    [XmlRoot("Template")]
    public class TemplateConfiguration
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateConfiguration"/> class.
        /// </summary>
        public TemplateConfiguration()
        {
            Conditions = new List<ConditionConfiguration>();
            Replacements = new List<ReplacementConfiguration>();
        }

        #endregion Constructors

        #region Public serialization properties

        /// <summary>
        /// Gets or sets the name of template.
        /// </summary>
        [XmlAttribute("Name")]
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the template name for header of block.
        /// </summary>
        [XmlAttribute("HeaderTemplate")]
        public string HeaderTemplate { get; set; }

        /// <summary>
        /// Gets or sets the file name of template.
        /// </summary>
        [XmlAttribute("FileName")]
        public string FileName { get; set; }

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

        /// <summary>
        /// Property used to serialize 'Conditions/Condition' xml node.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
        [XmlArray("Conditions")]
        [XmlArrayItem("Condition")]
        public List<ConditionConfiguration> Conditions { get; set; }

        /// <summary>
        /// Property used to serialize 'Replacements/Replacement' xml node.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
        [XmlArray("Replacements")]
        [XmlArrayItem("Replacement")]
        public List<ReplacementConfiguration> Replacements { get; set; }

        #endregion Public serialization properties
    }
}
