using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Service.Configuration.Model;
using AIT.TFS.SyncService.Service.Configuration.Serialization;

using ObjectQuery = AIT.TFS.SyncService.Contracts.Configuration.ObjectQuery;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// Class implements the functionality of configuration service - <see cref="IConfigurationService"/>.
    /// </summary>
    public class Configuration : IConfiguration
    {
        #region Fields
        private readonly Dictionary<string, MappingInfo> _mappings = new Dictionary<string, MappingInfo>();
        private readonly IList<IConfigurationItem> _configurations = new List<IConfigurationItem>();
        private readonly IList<IConfigurationItem> _headers = new List<IConfigurationItem>();
        private readonly IList<IObjectQuery> _objectQueries = new List<IObjectQuery>();

        private IList<IConfigurationFieldItem> _configFieldItems;
        private string _defaultMappingShowName;
        #endregion

        #region Implementation of IConfiguration

        /// <summary>
        /// Gets a value indicating whether the message for the specific problem should be showed.
        /// </summary>
        public bool ShowMappedCellNotExistsMessage { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the message for the specific problem should be showed.
        /// </summary>
        public bool ShowMappedFieldNotExistsMessage { get; private set; }

        /// <summary>
        /// Gets the option to synchronize stack ranks to the TFS.
        /// </summary>
        public bool UseStackRank { get; set; }

        /// <summary>
        /// Gets the list of customized buttons
        /// </summary>
        public List<IConfigurationButtonCustomization> Buttons { get; set; }

        /// <summary>
        /// Gets the test report configuration.
        /// </summary>
        public IConfigurationTest ConfigurationTest { get; private set; }

        /// <summary>
        /// Gets whether the Refresh ribbon button is enabled.
        /// </summary>
        public bool EnableRefresh { get; private set; }

        /// <summary>
        /// Gets whether the Publish ribbon button is enabled.
        /// </summary>
        public bool EnablePublish { get; private set; }

        /// <summary>
        /// Gets whether the Overview ribbon button is enabled.
        /// </summary>
        public bool EnableOverview
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets whether the checkbox that enables format ignoring is visible
        /// </summary>
        public bool EnableSwitchFormatting { get; private set; }

        /// <summary>
        /// Gets whether the checkbox that enables overwriting of conflicting items is visible
        /// </summary>
        public bool EnableSwitchConflictOverwrite { get; private set; }

        /// <summary>
        /// Gets whether the history comment option is enabled.
        /// </summary>
        public bool EnableHistoryComment { get; private set; }

        /// <summary>
        /// Gets whether the get work items option is enabled
        /// </summary>
        public bool EnableGetWorkItems { get; private set; }

        /// <summary>
        /// Gets whether the empty work item option is enabled
        /// </summary>
        public bool EnableEmpty { get; private set; }

        /// <summary>
        /// Gets whether the new work item option is enabled
        /// </summary>
        public bool EnableNew { get; private set; }

        /// <summary>
        /// Gets whether the edit default values option is enabled
        /// </summary>
        public bool EnableEditDefaultValues { get; private set; }

        /// <summary>
        /// Gets whether the delete ids option is enabled
        /// </summary>
        public bool EnableDeleteIds { get; private set; }

        /// <summary>
        /// Gets whether the area and iteration path option is enabled
        /// </summary>
        public bool EnableAreaIterationPath { get; private set; }

        /// <summary>
        /// Gets a value indicating whether during publish use plain text comparison for HTML field.
        /// Plain text comparison ignore font formatting, etc.
        /// </summary>
        public bool IgnoreFormatting { get; set; }

        /// <summary>
        /// Gets or sets whether publish operations should overwrite work item with revision conflicts
        /// </summary>
        public bool ConflictOverwrite { get; set; }

        /// <summary>
        /// Gets or sets whether the element should be showed in word
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public bool HideElementInWord
        {
            private get;
            set;
        }

        /// <summary>
        /// Gets the list of pre-operations.
        /// </summary>
        public IList<IConfigurationTestOperation> PreOperations { get; private set; }

        /// <summary>
        /// Gets the list of post-operations.
        /// </summary>
        public IList<IConfigurationTestOperation> PostOperations { get; private set; }

        /// <summary>
        /// Gets the AttachmentFolderMode
        /// </summary>
        public AttachmentFolderMode AttachmentFolderMode
        {
            get;
            set;
        }

        private string _typeOfHierarchyRelationships;
        /// <summary>
        /// Gets or sets the name of hierarchylevels
        /// This had to be adjusted, because of a typo in old mappings. 
        /// From now on the misspelled "TypeOfHierachyRelationships" and the correct one  "TypeOfHiera_r_chyRelationships" are supported
        /// </summary>
        public string TypeOfHierarchyRelationships
        {
            get
            {
                return _typeOfHierarchyRelationships;
            }
            set
            {
                if (_typeOfHierarchyRelationships != null)
                {
                    if (_typeOfHierarchyRelationships.Equals(""))
                    {
                        _typeOfHierarchyRelationships = value;
                    }

                }
                else
                {
                    _typeOfHierarchyRelationships = value;
                }

            }
        }

        /// <summary>
        /// Additional method to support als the misspelled version of "TypeOfHiera_r_chyRelationships"
        /// </summary>
        private string TypeOfHierachyRelationships
        {
            set
            {
                if (_typeOfHierarchyRelationships != null)
                {
                    if (_typeOfHierarchyRelationships.Equals(""))
                    {
                        _typeOfHierarchyRelationships = value;
                    }

                }
                else
                {
                    _typeOfHierarchyRelationships = value;
                }
            }
        }




        /// <summary>
        /// Method discovers all mapping files.
        /// </summary>
        /// <remarks>
        /// Method searches in folder '...\Local Settings\Application Data\AIT\WordToTFS\Template
        /// s' all files with wildcard '*.w2t'.
        /// </remarks>
        public void RefreshMappings()
        {
            ReadAllMappings();
        }

        /// <summary>
        /// Gets the folder where to read the mapping bundles from.
        /// </summary>
        public string MappingFolder
        {
            get
            {
                string templateDirectory =
                    Path.Combine(
                        Path.Combine(
                            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                         Constants.ApplicationCompany),
                            Constants.ApplicationName), Constants.MappingBundleSubfolder);
                Directory.CreateDirectory(templateDirectory);
                return templateDirectory;
            }
        }

        /// <summary>
        /// Gets the all mappings.
        /// </summary>
        public string[] Mappings
        {
            get { return _mappings.Keys.ToArray(); }
        }

        /// <summary>
        /// Gets the show name of default mapping bundle.
        /// </summary>
        public string DefaultMapping
        {
            get
            {
                string retVal = null;
                if (!string.IsNullOrEmpty(_defaultMappingShowName))
                    retVal = _defaultMappingShowName;
                if (string.IsNullOrEmpty(retVal) && Mappings.Length > 0)
                    retVal = Mappings[0];
                return retVal;
            }
        }

        /// <summary>
        /// Method checks if the specified mapping exists.
        /// </summary>
        /// <param name="showName">Show name of the mapping to activate.</param>
        /// <remarks>Show name is defined in the mapping xml file. (*.w2t)</remarks>
        public bool MappingExists(string showName)
        {
            return _mappings.Keys.Contains(showName);
        }

        /// <summary>
        /// Method activates for whole add in the given mapping and all associated.
        /// </summary>
        /// <param name="showName">Show name of the mapping to activate.</param>
        /// <remarks>Show name is defined in the mapping xml file. (*.w2t)</remarks>
        public void ActivateMapping(string showName)
        {
            if (_mappings.Count < 1)
            {
                RefreshMappings();
            }
            if (!MappingExists(showName))
            {
                // Show name not found. Probably the template bundle no more loaded.
                showName = DefaultMapping;
            }

            _configurations.Clear();
            _headers.Clear();
            ReadConfigurationItems(_mappings[showName].MappingFileName);
        }

        /// <summary>
        /// Activates all loaded mappings.
        /// </summary>
        public void ActivateAllMappings()
        {
            try
            {
                _configurations.Clear();
                _headers.Clear();
                foreach (var valuePair in _mappings)
                {
                    ReadConfigurationItems(valuePair.Value.MappingFileName);
                }
            }
            catch (Exception ex)
            {
                SyncServiceTrace.LogException(ex);
            }
        }

        /// <summary>
        /// Retrieves the configuration item which are needed for the mapping between the different objects.
        /// </summary>
        /// <returns>Method returns the <see cref="IConfigurationItem"/>.</returns>
        public IList<IConfigurationItem> GetConfigurationItems()
        {
            return _configurations;
        }

        /// <summary>
        /// Gets all headers in the configuration.
        /// </summary>
        public IList<IConfigurationItem> Headers
        {
            get { return _headers; }
        }


        /// <summary>
        /// Retries the field configuration of the specific configuration item.
        /// </summary>
        /// <returns>Method returns the list of <see cref="IConfigurationFieldItem"/>.</returns>
        public IList<IConfigurationFieldItem> GetFieldConfiguration()
        {
            if (_configFieldItems == null)
                _configFieldItems = new List<IConfigurationFieldItem>();
            return _configFieldItems;
        }

        /// <summary>
        /// Gets or sets the name of the current selected template.
        /// </summary>
        /// <value>
        /// The name of the current selected template.
        /// </value>
        /// <exception cref="System.NotImplementedException">
        /// </exception>
        public string CurrentMappingName
        {
            get;
            private set;
        }

        public IList<IVariable> Variables
        {
            get
            {
                //ReadConfigurationItems(_mappings[CurrentMappingName].MappingFileName);
                var fileName = _mappings[CurrentMappingName].MappingFileName;
                var serializer = new XmlSerializer(typeof(MappingConfiguration));
                MappingConfiguration mappingConfiguration;

                using (var stream = new StreamReader(fileName))
                {
                    mappingConfiguration = serializer.Deserialize(stream) as MappingConfiguration;
                }

                if (mappingConfiguration != null)
                {
                    var variables = mappingConfiguration.Variables.Cast<IVariable>().ToList();
                    return variables;
                }

                return null;
            }
        }

        /// <summary>
        /// Property used to show query tree collapsed or expanded.
        /// </summary>
        public bool CollapsQueryTree { get; set; }

        //private TestSpecReportingDefault _testSpecSpecReportDefault;
        //public ITestSpecReportDefault TestSpecReportDefaults
        //{
        //    get
        //    {
        //        return _testSpecSpecReportDefault;
        //    }
        //    set
        //    {
        //        if (null == _testSpecSpecReportDefault)
        //        {
        //            _testSpecSpecReportDefault = new TestSpecReportingDefault();
        //        }

        //        _testSpecSpecReportDefault.ReplaceValuesIfNotNull(value);
        //        //var newValue = value.TestSpecReport_CreateDocumentStructure;
        //        //var stayValue = _testSpecSpecReportDefault.TestSpecReport_CreateDocumentStructure;
        //        //_testSpecSpecReportDefault.TestSpecReport_CreateDocumentStructure = !string.IsNullOrEmpty(newValue) ? newValue : stayValue;

        //    }
        //}

        public IList<IObjectQuery> ObjectQueries
        {
            get
            {
                return _objectQueries;
            }
        }

        /// <summary>
        /// The server url for the autoconnect functionality
        /// </summary>
        public string DefaultServerUrl
        {
            get;
            private set;
        }

        /// <summary>
        /// The project name for the autoconnect functionality
        /// </summary>
        public string DefaultProjectName
        {
            get;
            private set;
        }

        /// <summary>
        /// Connect automaticly on the startup of Word
        /// </summary>
        public bool AutoConnect
        {
            get;
            private set;
        }

        /// <summary>
        /// Enable or disable the templatemanager
        /// </summary>
        public bool EnableTemplateManager
        {
            get;
            private set;
        }

        /// <summary>
        /// Enable or disable the selection of templates
        /// </summary>
        public bool EnableTemplateSelection
        {
            get;
            private set;
        }

        /// <summary>
        /// Determines whether to get only direct links only or recursive
        /// </summary>
        public bool GetDirectLinksOnly
        {
            get;
            private set;
        }

        /// <summary>
        /// Specifies a query that is automaticly refreshed during the start of WordToTFS
        /// </summary>
        public string AutoRefreshQuery
        {
            get;
            private set;
        }

        /// <summary>
        /// Reads the MappingList and deserializes all Mappings to ConfigurationItems.
        /// </summary>
        /// <param name="fileName">File to read.</param>
        public void ReadConfigurationItems(string fileName)
        {
            SyncServiceTrace.D("Read configuration items ...");
            SyncServiceTrace.I(Properties.Resources.LogService_LoadFile, fileName);

            if (!File.Exists(fileName))
            {
                SyncServiceTrace.W(Properties.Resources.LogService_LoadFile_FileDoesNotExist, fileName);
                return;
            }

            // Deserialize the xml file
            var serializer = new XmlSerializer(typeof(MappingConfiguration));
            MappingConfiguration mappingConfiguration;
            using (var stream = new StreamReader(fileName))
            {
                mappingConfiguration = serializer.Deserialize(stream) as MappingConfiguration;
            }
            CurrentMappingName = mappingConfiguration.ShowName;

            if (mappingConfiguration == null ||
                mappingConfiguration.MappingList == null ||
                mappingConfiguration.MappingList.Mappings == null ||
                mappingConfiguration.MappingList.Mappings.Length == 0)
            {
                ShowMappedCellNotExistsMessage = true;
                ShowMappedFieldNotExistsMessage = true;
                UseStackRank = false;
                ConfigurationTest = new ConfigurationTest();
                EnableRefresh = false;
                EnablePublish = true;
                EnableHistoryComment = false;
                EnableGetWorkItems = true;
                EnableOverview = true;
                EnableEmpty = true;
                EnableNew = true;
                EnableEditDefaultValues = true;
                EnableDeleteIds = true;
                EnableAreaIterationPath = true;
                IgnoreFormatting = false;
                HideElementInWord = false;
                TypeOfHierarchyRelationships = "";
                TypeOfHierachyRelationships = "";
                AttachmentFolderMode = AttachmentFolderMode.WithGuid;
                EnableTemplateManager = true;
                EnableTemplateSelection = true;
                GetDirectLinksOnly = false;
                AutoRefreshQuery = "";

                SyncServiceTrace.I(Properties.Resources.LogService_LoadFile_ResetToStandard, fileName);
                return;
            }

            PreOperations = null;
            PostOperations = null;
            Buttons = null;
            ShowMappedCellNotExistsMessage = mappingConfiguration.ShowMappedCellNotExistsMessage;
            ShowMappedFieldNotExistsMessage = mappingConfiguration.ShowMappedFieldNotExistsMessage;
            UseStackRank = mappingConfiguration.UseStackRank;
            ConfigurationTest = new ConfigurationTest(
                Path.GetDirectoryName(fileName),
                 mappingConfiguration.TestConfiguration);
            EnableRefresh = mappingConfiguration.EnableRefresh;
            EnablePublish = mappingConfiguration.EnablePublish;
            EnableSwitchConflictOverwrite = mappingConfiguration.EnableConflictOverwriteSwitch;
            EnableSwitchFormatting = mappingConfiguration.EnableIgnoreFormattingSwitch;
            EnableGetWorkItems = mappingConfiguration.EnableGetWorkItems;
            EnableOverview = mappingConfiguration.EnableOverview;
            EnableHistoryComment = mappingConfiguration.EnableHistoryComment;
            EnableEmpty = mappingConfiguration.EnableEmpty;
            EnableNew = mappingConfiguration.EnableNew;
            EnableEditDefaultValues = mappingConfiguration.EnableEditDefaultValues;
            EnableDeleteIds = mappingConfiguration.EnableDeleteIds;
            EnableAreaIterationPath = mappingConfiguration.EnableAreaIterationPath;
            IgnoreFormatting = mappingConfiguration.IgnoreFormatting;
            ConflictOverwrite = mappingConfiguration.ConflictOverwrite;
            HideElementInWord = mappingConfiguration.HideElementInWord;
            TypeOfHierarchyRelationships = mappingConfiguration.TypeOfHierarchyRelationships;
            TypeOfHierachyRelationships = mappingConfiguration.TypeOfHierachyRelationships;
            AttachmentFolderMode = mappingConfiguration.AttachmentFolderMode;
            EnableTemplateManager = mappingConfiguration.EnableTemplateManager;
            EnableTemplateSelection = mappingConfiguration.EnableTemplateSelection;
            GetDirectLinksOnly = mappingConfiguration.GetDirectLinksOnly;

            DefaultProjectName = mappingConfiguration.DefaultProjectName;
            DefaultServerUrl = mappingConfiguration.DefaultServerUrl;
            AutoConnect = mappingConfiguration.AutoConnect;
            CollapsQueryTree = mappingConfiguration.CollapsQueryTree;

            AutoRefreshQuery = mappingConfiguration.AutoRefreshQuery;

            if (mappingConfiguration.Buttons != null && mappingConfiguration.Buttons.Any())
            {
                Buttons = new List<IConfigurationButtonCustomization>();
                foreach (var button in mappingConfiguration.Buttons)
                {
                    Buttons.Add(button);
                }
            }

            if (mappingConfiguration.PreOperations != null && mappingConfiguration.PreOperations.Count > 0)
            {
                PreOperations = new List<IConfigurationTestOperation>();
                foreach (var operation in mappingConfiguration.PreOperations)
                    PreOperations.Add(new ConfigurationTestOperation(operation));
            }
            if (mappingConfiguration.PostOperations != null && mappingConfiguration.PostOperations.Count > 0)
            {
                PostOperations = new List<IConfigurationTestOperation>();
                foreach (var operation in mappingConfiguration.PostOperations)
                    PostOperations.Add(new ConfigurationTestOperation(operation));
            }

            //if (mappingConfiguration.TestSpecReportingDefault != null)
            //{
            //    TestSpecReportDefaults = mappingConfiguration.TestSpecReportingDefault;
            //}

            foreach (MappingElement mappingElement in mappingConfiguration.MappingList.Mappings)
            {
                var configurationItem = new ConfigurationItem(this, mappingElement, Path.GetDirectoryName(fileName));
                configurationItem.RelatedSchema = mappingConfiguration.ShowName;
                _configurations.Add(configurationItem);
            }

            foreach (var header in mappingConfiguration.Headers)
            {
                var configurationItem = new ConfigurationItem(this, header, Path.GetDirectoryName(fileName));
                configurationItem.RelatedSchema = mappingConfiguration.ShowName;
                _headers.Add(configurationItem);
            }

            foreach (var variable in mappingConfiguration.Variables)
            {

                //_variables.Add(mappingConfiguration.ShowName);
                //_variables.Add(mappingConfiguration.ShowName, variable);
            }

            foreach (var query in mappingConfiguration.ObjectQueries)
            {
                var queryExists = _objectQueries.Any(x => x.Name.Equals(query.Name));
                if (!queryExists)
                {
                    var objectQuery = new ObjectQuery();
                    objectQuery.Name = query.Name;

                    if (query.FilterOption != null)
                    {
                        IFilterOption filterOption = new FilterOption();
                        filterOption.Distinct = query.FilterOption.Distinct;
                        filterOption.Latest = query.FilterOption.Latest;
                        filterOption.FilterProperty = query.FilterOption.FilterProperty;
                        objectQuery.FilterOption = filterOption;
                    }

                    foreach (var dest in query.DestinationElements)
                    {
                        objectQuery.DestinationElements.Add(dest);
                    }

                    var filters = new WorkItemLinkFilter();

                    if (query.WorkItemLinkFilters != null)
                    {

                        filters.FilterType = query.WorkItemLinkFilters.FilterType;
                        foreach (var filter in query.WorkItemLinkFilters.Filters)
                        {
                            filters.Filters.Add(filter);
                        }

                        objectQuery.WorkItemLinkFilters = filters;
                    }

                    _objectQueries.Add(objectQuery);
                }
            }
        }

        /// <summary>
        /// Returns a <see cref="IConfigurationItem"/> of the <paramref name="workItemType"/> type.
        /// </summary>
        /// <param name="workItemType">Work item type.</param>
        /// <returns><see cref="IConfigurationItem"/> representation of the work item.</returns>
        public IConfigurationItem GetWorkItemConfiguration(string workItemType)
        {
            return _configurations.FirstOrDefault(x => x.WorkItemType.Equals(workItemType, StringComparison.OrdinalIgnoreCase));
        }


        /// <summary>
        /// Returns a <see cref="IConfigurationItem"/> representation of a work item type.
        /// </summary>
        /// <param name="workItemType">Type of the work item.</param>
        /// <param name="subtypeProvider">Function should evaluate property defined by input parameter and return the value of this property.</param>
        /// <returns><see cref="IConfigurationItem"/> representation of the <paramref name="workItemType"/>.</returns>
        public IConfigurationItem GetWorkItemConfigurationExtended(string workItemType, Func<string, string> subtypeProvider)
        {

            Guard.ThrowOnArgumentNull(subtypeProvider, "subtypeProvider");
            // find configuration with exact subtype matching
            var exactMatch =
                _configurations.FirstOrDefault(
                    x =>
                    string.IsNullOrEmpty(x.WorkItemSubtypeField) == false &&
                    string.IsNullOrEmpty(x.WorkItemSubtypeValue) == false &&
                    x.WorkItemType.Equals(workItemType, StringComparison.Ordinal) &&
                    x.WorkItemSubtypeValue.Equals(subtypeProvider(x.WorkItemSubtypeField), StringComparison.OrdinalIgnoreCase));

            // find configuration that does not care about subtypes
            var generalMatch =
                _configurations.FirstOrDefault(
                    x =>
                    string.IsNullOrEmpty(x.WorkItemSubtypeField) &&
                    string.IsNullOrEmpty(x.WorkItemSubtypeValue) &&
                    x.WorkItemType.Equals(workItemType, StringComparison.Ordinal));

            return exactMatch ?? generalMatch;
        }


        #endregion Implementation of IConfiguration

        #region Private methods
        /// <summary>
        /// Method finds all mapping files and checks the consistence.
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private void ReadAllMappings()
        {
            _defaultMappingShowName = null;
            _mappings.Clear();
            foreach (string subfolder in Directory.GetDirectories(MappingFolder))
            {
                // It is one time occured, that the folder in MappingFolder was deleted.
                // It takes some time and the method 'GetDirectories' returned this folder,
                // but the 'GetFiles' method was finished with exception folder not found.
                // Use exception handling
                try
                {
                    foreach (string mappingFile in Directory.GetFiles(subfolder, "*.w2t", SearchOption.AllDirectories))
                    {
                        SyncServiceTrace.I(Properties.Resources.LogService_LoadFile, mappingFile);

                        var serializer = new XmlSerializer(typeof(MappingConfigurationShort));
                        MappingConfigurationShort mappingConfiguration;
                        try
                        {
                            using (var stream = new StreamReader(mappingFile))
                            {
                                mappingConfiguration = serializer.Deserialize(stream) as MappingConfigurationShort;
                            }
                        }
                        catch (Exception ex)
                        {
                            SyncServiceTrace.LogException(ex);
                            mappingConfiguration = null;
                        }

                        if (mappingConfiguration == null || MappingExists(mappingConfiguration.ShowName))
                        {
                            continue;
                        }

                        var mi = new MappingInfo(mappingConfiguration.ShowName, mappingFile);
                        _mappings[mi.ShowName] = mi;
                        if (mappingConfiguration.DefaultMapping && string.IsNullOrEmpty(_defaultMappingShowName))
                        {
                            _defaultMappingShowName = mi.ShowName;
                        }
                    }
                }
                catch(Exception e)
                {
                    SyncServiceTrace.LogException(e);
                }
            }
            SyncServiceTrace.D("Mapping are read.");

        }


        #endregion
    }
}