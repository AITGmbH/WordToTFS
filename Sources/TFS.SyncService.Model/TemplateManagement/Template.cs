#region Usings
using AIT.TFS.SyncService.Contracts.TemplateManager;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.WindowModelBase;
using AIT.TFS.SyncService.Service.Configuration.Serialization;
using System;
using System.IO;
using System.Xml.Serialization;
#endregion

namespace AIT.TFS.SyncService.Model.TemplateManagement
{
    /// <summary>
    /// The class implements configuration of one template - w2t file with associated xml files.
    /// </summary>
    [XmlRoot("Template")]
    public class Template : ExtBaseModel, ITemplate
    {
        #region Consts
        private const string FaviconFileName = "favicon.ico";
        #endregion

        #region Private fields
        private string _showName;
        private TemplateState _templateState;
        private string _templateFile;
        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Template"/> class. Used in <see cref="XmlSerializer"/>.
        /// </summary>
        public Template()
        {
        }

        /// <summary>
        /// Initializes a new dummy instance of the <see cref="Template"/> class in case there was a deserialization error.
        /// </summary>
        /// <param name="templateFile">Template file - w2t file.</param>
        /// <param name="deserializationError">The error that occurred during deserialization.</param>
        public Template(string templateFile, Exception deserializationError)
        {
            Guard.ThrowOnArgumentNull(deserializationError, "deserializationError");
            TemplateFile = templateFile;

            // try to load the name of the template, even if the configuration could not be deserialized completely
            DetermineShowName();
            TemplateState = TemplateState.ErrorInTemplate;
            LoadError = deserializationError.Message;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Template"/> class.
        /// </summary>
        /// <param name="templateFile">Template file - w2t file.</param>
        /// <param name="templateConfiguration">The deserialized configuration from template file.</param>
        public Template(string templateFile, MappingConfiguration templateConfiguration)
        {
            Guard.ThrowOnArgumentNull(templateConfiguration, "templateConfiguration");

            TemplateFile = templateFile;
            ShowName = templateConfiguration.ShowName;
            TemplateState = TemplateState.Available;
        }
        #endregion

        #region Properies
        /// <summary>
        /// Gets or sets the name of the template.
        /// </summary>
        [XmlAttribute("ShowName")]
        public string ShowName
        {
            get
            {
                return _showName;
            }
            set
            {
                if (_showName != value)
                {
                    _showName = value;
                    TriggerPropertyChanged(nameof(ShowName));
                }
            }
        }

        /// <summary>
        /// Gets or sets the state of template - describes state of template.
        /// </summary>
        [XmlAttribute("TemplateState")]
        public TemplateState TemplateState
        {
            get
            {
                return _templateState;
            }
            set
            {
                if (_templateState != value)
                {
                    _templateState = value;
                    TriggerPropertyChanged(nameof(TemplateState));
                }
            }
        }

        /// <summary>
        /// Gets or sets the template file - w2t file.
        /// </summary>
        [XmlAttribute("TemplateFile")]
        public string TemplateFile
        {
            get { return _templateFile; }
            set
            {
                if (_templateFile != value)
                {
                    _templateFile = value;
                    TriggerPropertyChanged(nameof(TemplateFile));
                }
            }
        }

        /// <summary>
        /// Gets the error occurred on load time if any.
        /// </summary>
        [XmlIgnore]
        public string LoadError
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the name of the project this template is restricted to. This template should only be
        /// available if the user is connected to a project with the given name. If set to null or empty string
        /// the template is always available.
        /// </summary>
        [XmlIgnore]
        public string ProjectName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the name of the project collection this template is restricted to. This template should only be
        /// available if the user is connected to a project with the given name. If set to null or empty string
        /// the template is always available.
        /// </summary>
        [XmlIgnore]
        public string ProjectCollectionName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the name of the server this template is restricted to. This template should only be
        /// available if the user is connected to a server with the given name. If set to null or empty string
        /// the template is always available.
        /// </summary>
        [XmlIgnore]
        public string ServerName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets favicon.
        /// </summary>
        [XmlIgnore]
        public string TemplateFavicon
        {
            get { return Path.Combine(Path.GetDirectoryName(TemplateFile), FaviconFileName); }
        }

        #endregion

        #region Private methods
        /// <summary>
        /// Try to load a very simplified configuration to get the ShowName of the template.
        /// </summary>
        private void DetermineShowName()
        {
            var serializer = new XmlSerializer(typeof(MappingConfigurationShort));
            try
            {
                using (var stream = new StreamReader(TemplateFile))
                {
                    var mappingConfiguration = (MappingConfigurationShort) serializer.Deserialize(stream);
                    ShowName = mappingConfiguration.ShowName;
                }
            }
            catch (IOException)
            {
                ShowName = TemplateFile;
            }
            catch (InvalidOperationException)
            {
                ShowName = TemplateFile;
            }
        }
        #endregion
    }
}