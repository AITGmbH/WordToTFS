using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.Console
{
    /// <summary>
    /// Class used as top serialization class to serialize all settings
    /// </summary>
    [XmlRoot("Configuration", Namespace = "", IsNullable = false)]
    [XmlType("Configuration", Namespace = "")]
    public class Configuration
    {
        #region Public serialization properties 

        /// <summary>
        /// Gets or sets the document configuration part
        /// </summary>
        [XmlElement("Settings")]
        public Settings Settings { get; set; }

        /// <summary>
        /// Gets or sets the test configuration part.
        /// </summary>
        [XmlElement("TestConfiguration")]
        public TestConfiguration TestConfiguration { get; set; }

        #endregion Public serialization properties 
    }
}
