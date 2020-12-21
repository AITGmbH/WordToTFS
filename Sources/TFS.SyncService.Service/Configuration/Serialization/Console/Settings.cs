using System.Xml.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.Console
{

    /// <summary>
    /// The basic settings for the console extension.
    /// </summary>
    [XmlRoot("Settings", Namespace = "", IsNullable = false)]
    public class Settings
    {
        /// <summary>
        /// The server name
        /// </summary>
        [XmlAttribute("Server")]
        public string Server { get; set; }

        /// <summary>
        /// The project
        /// </summary>
        [XmlAttribute("Project")]
        public string Project { get; set; }

        /// <summary>
        /// The output file name
        /// </summary>

        [XmlAttribute("Filename")]
        public string Filename { get; set; }

        /// <summary>
        /// The W2T template that should be used
        /// </summary>
        [XmlAttribute("Template")]
        public string Template { get; set; }

        /// <summary>
        /// True to overwrite existing documents
        /// </summary>
        [XmlAttribute("Overwrite")]
        public bool Overwrite { get; set; }

        /// <summary>
        /// True to close after finishing
        /// </summary>
        [XmlAttribute("Close")]
        public bool CloseOnFinish { get; set; }

        /// <summary>
        /// True to hide word
        /// </summary>
        [XmlAttribute("WordHidden")]
        public bool WordHidden { get; set; }

        /// <summary>
        /// The path to a word template .dotx
        /// </summary>
        [XmlAttribute("DotxTemplate")]
        public string DotxTemplate { get; set; }


        /// <summary>
        /// The debug level for logging that should be used
        /// </summary>
        [XmlAttribute("DebugLevel")]
        public int DebugLevel { get; set; }
    }
}
