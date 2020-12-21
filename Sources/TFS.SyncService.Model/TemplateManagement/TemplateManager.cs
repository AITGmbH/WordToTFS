#region Usings
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.TemplateManager;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Properties;
#endregion

namespace AIT.TFS.SyncService.Model.TemplateManagement
{
    /// <summary>
    /// The class implements load of XML file containing configuration of more template bundles.
    /// </summary>
    public class TemplateManager
    {
        #region Consts
        private const string LocalTemplatesFileName = "Templates.xml";
        #endregion

        #region Private fields
        private readonly ObservableCollection<TemplateBundle> _templateBundles = new ObservableCollection<TemplateBundle>();
        private readonly ObservableCollection<Template> _availableTemplates = new ObservableCollection<Template>();
        #endregion

        #region Event Handlers
        /// <summary>
        /// Occurs when templates changed.
        /// </summary>
        public event EventHandler<EventArgs> TemplatesChanged;
        #endregion

        #region Contructor
        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateManager"/> class.
        /// Upgrades the templates and loads the template bundles.
        /// </summary>
        public TemplateManager()
        {
            UpgradeTemplateLocations();
            LoadPersistedTemplateBundles();
        }
        #endregion

        #region Public static methods
        /// <summary>
        /// Upgrade code: Template folder became TemplateCache so move all subfolders to TemplateCache as well as the templates.xml
        /// Deleting this upgrade code will require users to manually delete the template folder subfolders and add all their templates again
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private static void UpgradeTemplateLocations()
        {
            var oldLocalTemplatesFileLocation = Path.Combine(DefaultTemplateBundleLocation, LocalTemplatesFileName);

            // check if we already did this
            if (!File.Exists(oldLocalTemplatesFileLocation))
            {
                return;
            }

            try
            {
                if (!Directory.Exists(RootCacheDirectory))
                {
                    Directory.CreateDirectory(RootCacheDirectory);
                }

                foreach (var subDirectory in new DirectoryInfo(DefaultTemplateBundleLocation).GetDirectories())
                {
                    subDirectory.MoveTo(Path.Combine(RootCacheDirectory, subDirectory.Name));
                }

                File.Move(oldLocalTemplatesFileLocation, Path.Combine(RootCacheDirectory, LocalTemplatesFileName));
            }
            catch (Exception e)
            {
                SyncServiceTrace.LogException(e);
            }
        }
        #endregion

        #region Public properties
        /// <summary>
        /// Gets the folder with cached template bundles. Every template bundle is located in separate subfolder.
        /// </summary>
        public static string RootCacheDirectory
        {
            get
            {
                var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return Path.Combine(localAppData, Constants.ApplicationCompany, Constants.ApplicationName, Constants.MappingBundleSubfolder);
            }
        }

        /// <summary>
        /// Location of the default AIT template bundle in case it needs to be added again.
        /// </summary>
        public static string DefaultTemplateBundleLocation
        {
            get
            {
                var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return Path.Combine(localAppData, Constants.ApplicationCompany, Constants.ApplicationName, Constants.DefaultTemplateBundleSubfolder);
            }
        }

        /// <summary>
        /// Gets the configuration file where is stored configuration of all template bundles.
        /// </summary>
        public static string TemplateConfigurationFile
        {
            get
            {
                return Path.Combine(RootCacheDirectory, LocalTemplatesFileName);
            }
        }

        /// <summary>
        /// Gets a flat list of all templates for using in user interface.
        /// </summary>
        public IEnumerable<Template> AvailableTemplates
        {
            get
            {
                return _availableTemplates;
            }
        }

        /// <summary>
        /// Gets the collection of the persisted template bundles.
        /// </summary>
        public IEnumerable<TemplateBundle> TemplateBundles
        {
            get
            {
                return _templateBundles;
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// The method adds the template bundle to the collection of available template bundles
        /// and sorts the template bundles and templates in added template bundle.
        /// The properties <see cref="TemplateBundles"/> and <see cref="AvailableTemplates"/> are affected by this method.
        /// </summary>
        /// <param name="templateBundle">Template bundle to add.</param>
        public void AddTemplateBundle(TemplateBundle templateBundle)
        {
            if (templateBundle == null || TemplateBundles.Any(bundle => bundle.ShowName == templateBundle.ShowName))
            {
                // Already exists.
                return;
            }

            templateBundle.TemplatesChanged += HandleTemplateBundleTemplatesChangedEvent;
            _templateBundles.Add(templateBundle);

            Refresh();
        }

        private void Refresh()
        {
            SortTemplateBundles();
            FillAvailableTemplates();
            SaveTemplateBundles();
            RaiseTemplatesChangedEvent();
        }

        /// <summary>
        /// The method removes the template bundle from the collection of available template bundles.
        /// The properties <see cref="TemplateBundles"/> and <see cref="AvailableTemplates"/> are affected by this method.
        /// </summary>
        /// <param name="showName">Show name of template bundle to remove.</param>
        public void RemoveTemplateBundle(string showName)
        {
            var templateBundle = TemplateBundles.FirstOrDefault(bundle => bundle.ShowName == showName);
            if (templateBundle == null)
            {
                return;
            }

            templateBundle.TemplatesChanged -= HandleTemplateBundleTemplatesChangedEvent;
            _templateBundles.Remove(templateBundle);

            Refresh();

            // Call clean up after RaiseTemplatesChangedEvent to free favicon from template combo box
            templateBundle.Cleanup();
        }

        /// <summary>
        /// The method saves actually configured template bundles.
        /// </summary>
        public void SaveTemplateBundles()
        {
            try
            {
                var configuration = new TemplateConfiguration();
                foreach (var templateBundle in TemplateBundles)
                {
                    configuration.TemplateBundles.Add(templateBundle);
                }

                var serializer = new XmlSerializer(typeof(TemplateConfiguration));
                using (var streamWriter = new StreamWriter(TemplateConfigurationFile))
                {
                    serializer.Serialize(streamWriter, configuration);
                }
            }
            catch (InvalidOperationException ex)
            {
                SyncServiceTrace.LogException(ex);
            }
            catch (IOException ex)
            {
                SyncServiceTrace.LogException(ex);
            }
        }

        /// <summary>
        /// The method sorts the template bundles by its name.
        /// </summary>
        public void SortTemplateBundles()
        {
            var list = new List<TemplateBundle>(TemplateBundles);
            list.Sort((left, right) => string.Compare(left.ShowName, right.ShowName, true, CultureInfo.CurrentCulture));
            _templateBundles.Clear();
            foreach (var templateBundle in list)
            {
                _templateBundles.Add(templateBundle);
            }
        }

        /// <summary>
        /// Reload the template from its original source
        /// </summary>
        /// <param name="showName">Name of the template bundle to reload.</param>
        public void ReloadTemplateBundle(string showName)
        {
            var templateBundle = _templateBundles.FirstOrDefault(x => x.ShowName == showName);
            if (templateBundle == null)
            {
                return;
            }

            _templateBundles.Remove(templateBundle);
            var saveNeeded = ReloadTemplateBundle(templateBundle);

            SortTemplateBundles();
            FillAvailableTemplates();
            if (saveNeeded)
            {
                SaveTemplateBundles();
            }

            RaiseTemplatesChangedEvent();
        }

        #endregion

        #region Private methods
        /// <summary>
        /// The method fill ObservableCollection available templates with templates.
        /// </summary>
        private void FillAvailableTemplates()
        {
            _availableTemplates.Clear();
            foreach (var templateBundle in TemplateBundles)
            {
                foreach (var template in templateBundle.Templates)
                {
                    _availableTemplates.Add(template);
                }
            }
        }

        /// <summary>
        /// Loads the serialized template bundle information.
        /// </summary>
        /// <returns>The deserialized template configuration.</returns>
        private static TemplateConfiguration LoadTemplateBundleConfiguration()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(TemplateConfiguration));
                using (var streamReader = new StreamReader(TemplateConfigurationFile))
                {
                    return (TemplateConfiguration)serializer.Deserialize(streamReader);
                }
            }
            catch (InvalidOperationException ex)
            {
                SyncServiceTrace.LogException(ex);
            }
            catch (IOException ex)
            {
                SyncServiceTrace.LogException(ex);
            }

            return null;
        }

        /// <summary>
        /// Opens the serialize template bundle information and tries to load all sources into the cache.
        /// </summary>
        private void LoadPersistedTemplateBundles()
        {
            SyncServiceTrace.D(Resources.LoadTemplateInfo, TemplateConfigurationFile);
            var saveNeeded = false;
            if (File.Exists(TemplateConfigurationFile))
            {
                var configuration = LoadTemplateBundleConfiguration();
                if (configuration != null)
                {
                    foreach (var templateBundle in configuration.TemplateBundles)
                    {
                        saveNeeded = ReloadTemplateBundle(templateBundle);
                    }
                }
            }

            SyncServiceTrace.D(Resources.LoadTemplateBundles, _templateBundles.Count);
            if (_templateBundles.Count == 0)
            {
                saveNeeded = true;
                var templateBundle = CreateDefaultTemplateBundle();
                templateBundle.TemplatesChanged += HandleTemplateBundleTemplatesChangedEvent;
                _templateBundles.Add(templateBundle);
            }

            SortTemplateBundles();
            FillAvailableTemplates();
            if (saveNeeded)
            {
                SaveTemplateBundles();
            }
        }

        /// <summary>
        /// Loads a template bundle from its original source and applies stored settings.
        /// </summary>
        /// <param name="templateBundle">The template bundle to reload.</param>
        /// <returns>True if the original location is available, false otherwise.</returns>
        private bool ReloadTemplateBundle(TemplateBundle templateBundle)
        {
            // Load the actual source if its available, otherwise load the cached version
            var reloadedTemplateBundle = CreateTemplateBundle(templateBundle.ShowName, templateBundle.SourceLocation, templateBundle.Guid);
            var isCached = (reloadedTemplateBundle.TemplateBundleState & TemplateBundleStates.TemplateBundleCached) == TemplateBundleStates.TemplateBundleCached;
            var isAvailable = (reloadedTemplateBundle.TemplateBundleState & TemplateBundleStates.TemplateBundleAvailable) == TemplateBundleStates.TemplateBundleAvailable;

            if (isAvailable)
            {
                // merge old settings into the reloaded template bundle
                reloadedTemplateBundle.HasProjectMappedFolderHierarchy = templateBundle.HasProjectMappedFolderHierarchy;
                foreach (var template in templateBundle.Templates)
                {
                    var reloadedTemplate = reloadedTemplateBundle.Templates.FirstOrDefault(x => x.ShowName.Equals(template.ShowName, StringComparison.OrdinalIgnoreCase));
                    if (reloadedTemplate != null &&
                        reloadedTemplate.TemplateState != TemplateState.ErrorInTemplate &&
                        template.TemplateState == TemplateState.Disabled)
                    {
                        reloadedTemplate.TemplateState = TemplateState.Disabled;
                    }
                }

                LoadTemplateBundle(reloadedTemplateBundle);
            }
            else if (isCached)
            {
                SyncServiceTrace.I(Resources.TemplateBundleNotAvailable, templateBundle.ShowName, templateBundle.SourceLocation);
                LoadTemplateBundle(templateBundle);
                templateBundle.TemplateBundleState |= TemplateBundleStates.TemplateBundleCached;
            }

            return isAvailable;
        }

        /// <summary>
        /// Adds a template bundle to the list of available template bundles.
        /// </summary>
        private void LoadTemplateBundle(TemplateBundle templateBundle)
        {
            templateBundle.SetTemplateBundleState();
            templateBundle.TemplatesChanged += HandleTemplateBundleTemplatesChangedEvent;
            _templateBundles.Add(templateBundle);

            // toggle the bool value twice to trigger the value changed event in the setter.
            // The value changes templates which might not have been deserialized when the value is set.
            templateBundle.HasProjectMappedFolderHierarchy = !templateBundle.HasProjectMappedFolderHierarchy;
            templateBundle.HasProjectMappedFolderHierarchy = !templateBundle.HasProjectMappedFolderHierarchy;
        }

        /// <summary>
        /// The method creates the default template bundle provided by WordToTFS
        /// </summary>
        private static TemplateBundle CreateDefaultTemplateBundle()
        {
            return CreateTemplateBundle(Resources.DefaultTemplateBundleName, DefaultTemplateBundleLocation, Guid.NewGuid());
        }

        /// <summary>
        /// The method creates the template bundle for provided name and path.
        /// </summary>
        /// <param name="showName">Show name of the template bundle. Provided by the user in user interface.</param>
        /// <param name="sourceLocation">Path of the template bundle. Provided by the user in user interface.</param>
        /// <param name="templateBundleGuid">Globally unique identifier of template bundle</param>
        /// <returns>Created template bundle.</returns>
        private static TemplateBundle CreateTemplateBundle(string showName, string sourceLocation, Guid templateBundleGuid)
        {
            var templateBundle = new TemplateBundle(showName, sourceLocation, templateBundleGuid);
            templateBundle.SortTemplates();
            return templateBundle;
        }

        /// <summary>
        /// Handles <see cref="TemplateBundle.TemplatesChanged"/> event.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">EventData</param>
        private void HandleTemplateBundleTemplatesChangedEvent(object sender, EventArgs e)
        {
            SaveTemplateBundles();
            RaiseTemplatesChangedEvent();
        }

        /// <summary>
        /// The method calls the event <see cref="TemplatesChanged"/>.
        /// </summary>
        private void RaiseTemplatesChangedEvent()
        {
            if (TemplatesChanged != null)
            {
                SyncServiceTrace.D(Resources.TemplatesChanged);
                TemplatesChanged(this, EventArgs.Empty);
            }
        }

        #endregion
    }
}
