using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;

namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Interface defines the functionality of configuration service.
    /// </summary>
    public interface IConfiguration
    {
        /// <summary>
        /// Gets the test report configuration.
        /// </summary>
        IConfigurationTest ConfigurationTest { get; }

        /// <summary>
        /// Gets a value indicating whether the message for the specific problem should be showed.
        /// </summary>
        bool ShowMappedCellNotExistsMessage { get; }

        /// <summary>
        /// Gets a value indicating whether the message for the specific problem should be showed.
        /// </summary>
        bool ShowMappedFieldNotExistsMessage { get; }

        /// <summary>
        /// Gets the option to synchronize stack ranks to the TFS.
        /// </summary>
        bool UseStackRank
        {
            get;
            set;
        }


        /// <summary>
        /// Name of the hierarchy relationships
        /// </summary>
        string TypeOfHierarchyRelationships { get; }

        /// <summary>
        /// Gets whether the Refresh ribbon button is enabled.
        /// </summary>
        bool EnableRefresh { get; }

        /// <summary>
        /// Gets whether the Publish ribbon button is enabled.
        /// </summary>
        bool EnablePublish { get; }

        /// <summary>
        /// Gets whether the Overview ribbon button is enabled.
        /// </summary>
        bool EnableOverview { get; }

        /// <summary>
        /// Gets whether the checkbox that enables format ignoring is visible
        /// </summary>
        bool EnableSwitchFormatting { get; }

        /// <summary>
        /// Gets whether the checkbox that enables overwriting of conflicting items is visible
        /// </summary>
        bool EnableSwitchConflictOverwrite { get; }

        /// <summary>
        /// Gets whether the Get Work Items ribbon button is enabled
        /// </summary>
        bool EnableGetWorkItems { get; }

        /// <summary>
        /// Gets whether the history comment option is enabled.
        /// </summary>
        bool EnableHistoryComment { get; }

        /// <summary>
        /// Gets whether the empty work item option is enabled
        /// </summary>
        bool EnableEmpty { get; }

        /// <summary>
        /// Gets whether the new work item option is enabled
        /// </summary>
        bool EnableNew { get; }

        /// <summary>
        /// Gets whether the edit default values option is enabled
        /// </summary>
        bool EnableEditDefaultValues { get; }

        /// <summary>
        /// Gets whether the delete ids option is enabled
        /// </summary>
        bool EnableDeleteIds { get; }

        /// <summary>
        /// Gets whether the area and iteration path option is enabled
        /// </summary>
        bool EnableAreaIterationPath { get; }

        /// <summary>
        /// Gets a value indicating whether during publish use plain text comparison for HTML field. 
        /// Plain text comparison ignore font formatting, etc.
        /// </summary>
        bool IgnoreFormatting { get; set; }

        /// <summary>
        /// Gets or sets whether publish operations should overwrite work item with revision conflicts
        /// </summary>
        bool ConflictOverwrite { get; set; }

        /// <summary>
        /// Gets the folder where to read the mapping bundles from.
        /// </summary>
        string MappingFolder { get; }

        /// <summary>
        /// Gets the list of pre-operations.
        /// </summary>
        IList<IConfigurationTestOperation> PreOperations { get; }

        /// <summary>
        /// Gets the list of post-operations.
        /// </summary>
        IList<IConfigurationTestOperation> PostOperations { get; }

        ///// <summary>
        ///// Gets the list of post-operations.
        ///// </summary>l
        List<IConfigurationButtonCustomization> Buttons { get; set; }

        /// <summary>
        /// Gets the all mappings.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        string[] Mappings { get; }



        /// <summary>
        /// Gets the show name of default mapping bundle.
        /// </summary>
        string DefaultMapping { get; }

        /// <summary>
        /// Method discovers all mapping files.
        /// </summary>
        /// <remarks>
        /// Method searches in folder '...\Local Settings\Application Data\AIT\WordToTFS\Templates' all files with wildcard '*.w2t'.
        /// </remarks>
        void RefreshMappings();

        /// <summary>
        /// Method checks if the specified mapping exists.
        /// </summary>
        /// <param name="showName">Show name of the mapping to activate.</param>
        /// <remarks>Show name is defined in the mapping xml file. (*.w2t)</remarks>
        bool MappingExists(string showName);

        /// <summary>
        /// Method activates for whole add in the given mapping and all associated
        /// </summary>
        /// <param name="showName">Show name of the mapping to activate.</param>
        /// <remarks>Show name is defined in the mapping xml file (*.w2t).</remarks>
        void ActivateMapping(string showName);

        /// <summary>
        /// Retrieves the configuration item which are needed for the mapping between the different objects.
        /// </summary>
        /// <returns>Method returns the <see cref="IConfigurationItem"/>.</returns>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")]
        IList<IConfigurationItem> GetConfigurationItems();

        /// <summary>
        /// Retries the field configuration of the specific configuration item.
        /// </summary>
        /// <returns>List of <see cref="IConfigurationFieldItem"/>.</returns>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")]
        IList<IConfigurationFieldItem> GetFieldConfiguration();

        /// <summary>
        /// Gets all headers in the configuration.
        /// </summary>
        IList<IConfigurationItem> Headers { get; }

        /// <summary>
        /// Read the configuration items from a *.w2t file.
        /// </summary>
        /// <param name="fileName">File name of the file.</param>
        void ReadConfigurationItems(string fileName);

        /// <summary>
        /// Returns a <see cref="IConfigurationItem"/> representation of a work item type.
        /// </summary>
        /// <param name="workItemType">Type of the work item.</param>
        /// <returns><see cref="IConfigurationItem"/> representation of the <paramref name="workItemType"/>.</returns>
        IConfigurationItem GetWorkItemConfiguration(string workItemType);

        /// <summary>
        /// Returns a <see cref="IConfigurationItem"/> representation of a work item type.
        /// </summary>
        /// <param name="workItemType">Type of the work item.</param>
        /// <param name="subtypeProvider">Function should evaluate property defined by input parameter and return the value of this property.</param>
        /// <returns><see cref="IConfigurationItem"/> representation of the <paramref name="workItemType"/>.</returns>
        IConfigurationItem GetWorkItemConfigurationExtended(string workItemType, Func<string, string> subtypeProvider);

        /// <summary>
        /// Activates all mappings.
        /// </summary>
        void ActivateAllMappings();

        /// <summary>
        /// Gets the AttachmentFolderMode
        /// </summary>
        AttachmentFolderMode AttachmentFolderMode { get; }

        /// <summary>
        /// The List of ObjectQueries for the Configuration
        /// </summary>
        IList<IObjectQuery> ObjectQueries
        {
            get;
        }

        /// <summary>
        /// The server url for the autoconnect functionality
        /// </summary>
        string DefaultServerUrl
        {
            get;
        }

        /// <summary>
        /// The project name for the autoconnect functionality
        /// </summary>
        string DefaultProjectName
        {
            get;
        }

        /// <summary>
        /// Connect automaticly on start on word
        /// </summary>
        bool AutoConnect
        {
            get;
     
        }

        /// <summary>
        /// Enable or disable the templatemanager
        /// </summary>
        bool EnableTemplateManager
        {
            get;
           
        }

        /// <summary>
        /// Enable or disable the selection of templates
        /// </summary>
        bool EnableTemplateSelection
        {
            get;

        }
        /// <summary>
        /// Determines whether to get direct links only or recursive
        /// </summary>
        bool GetDirectLinksOnly
        {
            get;
        }

        /// <summary>
        /// Specifies a query that is automaticly refreshed during the start of WordToTFS
        /// </summary>
        string AutoRefreshQuery
        {
            get;
        }

        /// <summary>
        /// Gets the name of the current mapping.
        /// </summary>
        /// <value>
        /// The name of the current mapping.
        /// </value>
        string CurrentMappingName
        {
            get;
        }

        /// <summary>
        /// Gets the variables.
        /// </summary>
        /// <value>
        /// The variables.
        /// </value>
        IList<IVariable> Variables
        {
            get;
        }

        /// <summary>
        /// Property used to show query tree collapsed or expanded.
        /// </summary>
        bool CollapsQueryTree { get; set; }
    }
}