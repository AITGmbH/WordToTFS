#region Usings
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.TemplateManager;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Properties;
using AIT.TFS.SyncService.Model.WindowModelBase;
using AIT.TFS.SyncService.Service.Configuration.Serialization;
#endregion

namespace AIT.TFS.SyncService.Model.TemplateManagement
{
    /// <summary>
    /// The class implements configuration of one template bundle - folder with available w2t templates.
    /// </summary>
    [XmlRoot("TemplateBundle")]
    public class TemplateBundle : ExtBaseModel, ITemplateBundle
    {
        #region Private fields
        private string _showName;
        private TemplateBundleType _type;
        private TemplateBundleStates _state;
        private bool _hasProjectMappedFolderHierarchy;
        private Guid _guid;
        private string _sourceLocation;
        private DateTime _lastUpdate;
        private List<Template> _templates;
        #endregion

        #region Event Handlers
        /// <summary>
        /// Event occurs if any change in any template bundle is committed.
        /// </summary>
        public event EventHandler<EventArgs> TemplatesChanged;
        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateBundle"/> class. Used in <see cref="XmlSerializer"/>.
        /// </summary>
        public TemplateBundle()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateBundle"/> class.
        /// </summary>
        /// <param name="showName">The show name of template bundle.</param>
        /// <param name="sourceLocation">Source location of template bundle.</param>
        /// <param name="templateBundleIdentifier">Globally unique identifier of the template bundle.</param>
        public TemplateBundle(string showName, string sourceLocation, Guid templateBundleIdentifier)
        {
            Guard.ThrowOnArgumentNull(showName, "showName");
            Guard.ThrowOnArgumentNull(sourceLocation, "sourceLocation");

            if (templateBundleIdentifier == null)
            {
                throw new ArgumentNullException("templateBundleIdentifier");
            }

            Templates = new List<Template>();
            ShowName = showName;
            SourceLocation = sourceLocation;
            Guid = templateBundleIdentifier;
            DetermineTemplates();
            HookUpEvents();
        }

        #endregion Constructors

        #region Public serialization properties

        /// <summary>
        /// Gets or sets the name of the template bundle.
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
        /// Gets or sets the origin location of the template bundle.
        /// </summary>
        [XmlAttribute("SourceLocation")]
        public string SourceLocation
        {
            get
            {
                return _sourceLocation;
            }
            set
            {
                if (_sourceLocation != value)
                {
                    _sourceLocation = value;
                    TriggerPropertyChanged(nameof(SourceLocation));
                }
            }
        }

        /// <summary>
        /// Gets or sets the globally unique identifier of the template bundle.
        /// Globally unique identifier is the name of subfolder too, where is the copy of template bundle stored.
        /// </summary>
        [XmlAttribute("TemplateBundleGuid")]
        public Guid Guid
        {
            get
            {
                return _guid;
            }
            set
            {
                if (_guid != value)
                {
                    _guid = value;
                    TriggerPropertyChanged(nameof(Guid));
                }
            }
        }

        /// <summary>
        /// Gets or sets the type of template bundle - type defines the origin of the template bundle.
        /// </summary>
        [XmlAttribute("TemplateBundleType")]
        public TemplateBundleType TemplateBundleType
        {
            get
            {
                return _type;
            }
            set
            {
                if (_type != value)
                {
                    _type = value;
                    TriggerPropertyChanged(nameof(TemplateBundleType));
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether this template bundle has a project mapped folder hierarchy.
        /// Templates within project mapped folder hierarchy should only be available if the user is
        /// connected to the project with the same name.
        /// </summary>
        /// <value>
        /// <c>true</c> if this template bundle has a project mapped folder hierarchy; otherwise, <c>false</c>.
        /// </value>
        [XmlAttribute("HasProjectMappedFolderHierarchy")]
        public bool HasProjectMappedFolderHierarchy
        {
            get
            {
                return _hasProjectMappedFolderHierarchy;
            }
            set
            {
                if (_hasProjectMappedFolderHierarchy != value)
                {
                    _hasProjectMappedFolderHierarchy = value;
                    TriggerPropertyChanged(nameof(HasProjectMappedFolderHierarchy));

                    foreach (var template in Templates)
                    {
                        var parentDirectory = Directory.GetParent(template.TemplateFile);
                        if (parentDirectory != null)
                        {
                            template.ProjectName = value ? parentDirectory.Name : null;

                            if (parentDirectory.Parent != null && !parentDirectory.Parent.Name.Equals(Guid.ToString()))
                            {
                                template.ProjectCollectionName = value ? parentDirectory.Parent.Name : null;

                                if (parentDirectory.Parent.Parent != null && !parentDirectory.Parent.Parent.Name.Equals(Guid.ToString()))
                                {
                                    template.ServerName = value ? parentDirectory.Parent.Parent.Name : null;
                                }

                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the state of whole template bundle - describes state of whole template bundle.
        /// </summary>
        [XmlAttribute("TemplateBundleState")]
        public TemplateBundleStates TemplateBundleState
        {
            get
            {
                return _state;
            }
            set
            {
                if (_state != value)
                {
                    _state = value;
                    TriggerPropertyChanged(nameof(TemplateBundleState));
                }
            }
        }

        /// <summary>
        /// Gets or sets the time last updated.
        /// </summary>
        /// <value>The last update.</value>
        [XmlAttribute("LastUpdate")]
        public DateTime LastUpdate
        {
            get
            {
                return _lastUpdate;
            }
            set
            {
                if (_lastUpdate != value)
                {
                    _lastUpdate = value;
                    TriggerPropertyChanged(nameof(LastUpdate));
                }
            }
        }

        /// <summary>
        /// Gets or sets the templates of the template bundle.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "We need sort functionality that a collection does not offer."), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "XML serializer cannot deserialize readonly properties.")]
        [XmlArray("Templates")]
        [XmlArrayItem("Template")]
        public List<Template> Templates
        {
            get
            {
                return _templates;
            }
            set
            {
                if (_templates != null)
                {
                    foreach (var template in _templates)
                        template.PropertyChanged -= HandleTemplatePropertyChangedEvent;
                }
                _templates = value;
                HookUpEvents();
            }
        }

        #endregion Public serialization properties

        #region Public properties

        /// <summary>
        /// Gets a value indicating whether the template bundle is removable.
        /// </summary>
        [XmlIgnore]
        public bool Removable
        {
            get
            {
                return ShowName != Resources.DefaultTemplateBundleName;
            }
        }

        /// <summary>
        /// Gets or sets the folder of cached template bundle.
        /// </summary>
        [XmlIgnore]
        public string TemplateBundleCacheLocation
        {
            get
            {
                // If this property is called. Guid must be already presented.
                Debug.Assert(Guid != Guid.Empty, "Logic problem in source. Check the implementation!");
                return Path.Combine(TemplateManager.RootCacheDirectory, Guid.ToString());
            }
        }

        #endregion Public properties

        #region Public methods

        /// <summary>
        /// The method sets the state of template bundle regarding on templates.
        /// </summary>
        public void SetTemplateBundleState()
        {
            if (Templates.Any(template => template.TemplateState == TemplateState.ErrorInTemplate))
            {
                TemplateBundleState |= TemplateBundleStates.AtLeastOneTemplateIsFaulty;
            }
            else
            {
                TemplateBundleState &= ~TemplateBundleStates.AtLeastOneTemplateIsFaulty;
            }
            if (Templates.Any(template => template.TemplateState == TemplateState.Disabled))
            {
                TemplateBundleState |= TemplateBundleStates.AtLeastOneTemplateDisabled;
            }
            else
            {
                TemplateBundleState &= ~TemplateBundleStates.AtLeastOneTemplateDisabled;
            }
        }

        /// <summary>
        /// The method sorts the templates.
        /// </summary>
        public void SortTemplates()
        {
            Templates.Sort((left, right) => string.Compare(left.ShowName, right.ShowName, true, CultureInfo.CurrentCulture));
            HookUpEvents();
        }

        /// <summary>
        /// The method deletes all cached files.
        /// </summary>
        public void Cleanup()
        {
            try
            {
                Directory.Delete(TemplateBundleCacheLocation, true);
            }
            catch (IOException e)
            {
                SyncServiceTrace.LogException(e);
            }
            catch (UnauthorizedAccessException e)
            {
                SyncServiceTrace.LogException(e);
            }

        }

        #endregion Public methods

        #region Private methods

        /// <summary>
        /// The method determines all available templates for the template bundle.
        /// It is <see cref="Path"/> used.
        /// </summary>
        private void DetermineTemplates()
        {
            SetTemplateBundleType();
            LastUpdate = DateTime.Now;
            CollectTemplates();
            SetTemplateBundleState();
        }

        /// <summary>
        /// The method determines source type of the template bundle.
        /// </summary>
        private void SetTemplateBundleType()
        {
            var local = new Regex(@"^\w{1}:\\");
            var http = new Regex(@"^(http|https)://");
            var www = new Regex(@"^w{3}\..+\.\w{2,}");
            var unc = new Regex(@"^\\\\");

            if (local.IsMatch(_sourceLocation))
            {
                TemplateBundleType = TemplateBundleType.LocalBundle;
            }
            else if (http.IsMatch(_sourceLocation))
            {
                TemplateBundleType = TemplateBundleType.WebBundle;
            }
            else if (unc.IsMatch(_sourceLocation))
            {
                TemplateBundleType = TemplateBundleType.UncBundle;
            }
            else if (www.IsMatch(_sourceLocation))
            {
                _sourceLocation = string.Concat("http://", _sourceLocation);
                TemplateBundleType = TemplateBundleType.WebBundle;
            }
            else
            {
                TemplateBundleType = TemplateBundleType.Unknown;
            }
        }

        /// <summary>
        /// The method checks availability of the template bundle and copies it to the cache if it is available.
        /// </summary>
        private void CollectTemplates()
        {
            TemplateBundleState = TemplateBundleStates.None;

            switch (TemplateBundleType)
            {
                case TemplateBundleType.LocalBundle:
                case TemplateBundleType.UncBundle:

                    // Set to cached if source does not exist but files are cached
                    if (!Directory.Exists(SourceLocation))
                    {
                        if (Directory.Exists(TemplateBundleCacheLocation))
                        {
                            TemplateBundleState = TemplateBundleStates.TemplateBundleCached;
                        }
                    }
                    else
                    {
                        // The source location is available.
                        // We need to copy the source files every time.
                        try
                        {
                            CopyLocalSourceToCache();
                            TemplateBundleState = TemplateBundleStates.TemplateBundleAvailable;
                        }
                        catch (IOException e)
                        {
                            SyncServiceTrace.LogException(e);
                        }
                        catch (UnauthorizedAccessException e)
                        {
                            SyncServiceTrace.LogException(e);
                        }
                    }
                    break;
                case TemplateBundleType.WebBundle:
                    Debug.Assert(false, "Not implemented yet");
                    break;
            }
        }

        /// <summary>
        /// The method hooks up to the <see cref="INotifyPropertyChanged.PropertyChanged"/> event in every template.
        /// </summary>
        private void HookUpEvents()
        {
            foreach (var template in Templates)
            {
                template.PropertyChanged += HandleTemplatePropertyChangedEvent;
            }
        }

        #endregion Private methods

        #region Private methods - copy file related

        private void CopyLocalSourceToCache()
        {
            if (Directory.Exists(TemplateBundleCacheLocation))
            {
                Directory.Delete(TemplateBundleCacheLocation, true);
            }

            Directory.CreateDirectory(TemplateBundleCacheLocation);

            CopyFilesRecursively(new DirectoryInfo(SourceLocation), new DirectoryInfo(TemplateBundleCacheLocation));
        }


        private void CopyFilesRecursively(DirectoryInfo source, DirectoryInfo target)
        {
            LastUpdate = DateTime.Now;
            var serializer = new XmlSerializer(typeof(MappingConfiguration));

            // Recursively search subfolders
            foreach (var directory in source.GetDirectories())
            {
                CopyFilesRecursively(directory, target.CreateSubdirectory(directory.Name));
            }

            // Copy all files to cache
            foreach (FileInfo file in source.GetFiles())
            {
                var cachedFilename = Path.Combine(target.FullName, file.Name);
                file.CopyTo(cachedFilename, true);
                File.SetAttributes(cachedFilename, FileAttributes.Normal);

                // if we found a configuration file, load it
                if (file.Extension.Equals(".w2t"))
                {
                    try
                    {
                        using (var stream = new StreamReader(cachedFilename))
                        {
                            var mappingConfiguration = (MappingConfiguration) serializer.Deserialize(stream);
                            var template = new Template(cachedFilename, mappingConfiguration);
                            Templates.Add(template);
                        }
                    }
                    catch (InvalidOperationException ex)
                    {
                        Templates.Add(new Template(cachedFilename, ex));
                        SyncServiceTrace.LogException(ex);
                    }
                }
            }
        }

        #endregion Private methods - copy file related

        /// <summary>
        /// Handles <see cref="INotifyPropertyChanged.PropertyChanged"/> event.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event data.</param>
        private void HandleTemplatePropertyChangedEvent(object sender, PropertyChangedEventArgs e)
        {
            LastUpdate = DateTime.Now;
            RaiseTemplatesChangedEvent();
            SetTemplateBundleState();
        }

        /// <summary>
        /// The method calls the event <see cref="TemplatesChanged"/>.
        /// </summary>
        private void RaiseTemplatesChangedEvent()
        {
            if (TemplatesChanged != null)
            {
                TemplatesChanged(this, EventArgs.Empty);
            }
        }
    }
}
