namespace AIT.TFS.SyncService.Model.TemplateManagement
{
    #region Usings
    using System.Collections.ObjectModel;
    using System.Xml.Serialization;
    #endregion

    /// <summary>
    /// The class implements configuration of all template bundles.
    /// </summary>
    [XmlRoot("WordToTFSConfiguration")]
    public class TemplateConfiguration
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateConfiguration"/> class.
        /// </summary>
        public TemplateConfiguration()
        {
            TemplateBundles = new ObservableCollection<TemplateBundle>();
        }

        /// <summary>
        /// Gets or sets the templates of the template bundle.
        /// </summary>
        [XmlArray("TemplateBundles"), XmlArrayItem("TemplateBundle")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "XML serializer cannot deserialize readonly properties.")]
        public ObservableCollection<TemplateBundle> TemplateBundles { get; set; }
    }
}
