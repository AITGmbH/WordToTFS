namespace AIT.TFS.SyncService.Contracts.WorkItemObjects
{
    using System;
    using System.Collections.Generic;
    using Configuration;
    using Exceptions;
    using WorkItemCollections;
    using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;
    using Microsoft.TeamFoundation.ProcessConfiguration.Client; 
       
    /// <summary>
    /// Generic interface to access work items from different sources
    /// </summary>
    public interface IWorkItemSyncAdapter
    {
        /// <summary>
        /// A collection of all opened work items. The collection will be initialized when calling open
        /// </summary>
        IWorkItemCollection WorkItems { get; }

        /// <summary>
        /// A collection of work item colors.
        /// </summary>
        IDictionary<string , WorkItemTypeAppearance> WorkItemColors { get; }

        /// <summary>
        /// Indicates whether this adapter has been opened. The list of work items will not be initialized unless
        /// an adapter is opened
        /// </summary>
        bool IsOpen { get; }

        /// <summary>
        /// Open the adapter for read and write of work items
        /// </summary>
        /// <param name="ids">Ids of the work items to open</param>
        /// <returns>Returns true on success, false otherwise</returns>
        bool Open(int[] ids);

        /// <summary>
        /// Connects to the TFS adapter a queries the requested work items. Use this method if you want open each
        /// configured in the link format syntax)
        /// </summary>
        /// <returns>Returns true on success, false otherwise</returns>
        bool OpenWithConfigurations(Dictionary<int, IConfigurationItem> configurations );

        /// <summary>
        /// Open the adapter for read and write of work items. Use this method if you need to query more fields
        /// than only the fields specified in the work item mapping. (Link expanding uses this to query additional fields
        /// configured in the link format syntax)
        /// </summary>
        /// <param name="ids">Ids of the work items to open</param>
        /// <param name="fields">list of fields to query</param>
        /// <returns>Returns true on success, false otherwise</returns>
        bool Open(int[] ids, IList<string> fields);

        /// <summary>
        /// Open the adapter for read and write of work items. Use this method if you need to query more fields
        /// than only the fields specified in the work item mapping. (Link expanding uses this to query additional fields
        /// configured in the link format syntax)
        /// </summary>
        /// <param name="linkItem">Link Item</param>
        /// <param name="fields">list of fields to query</param>
        /// <returns>Returns true on success, false otherwise</returns>
        bool Open(KeyValuePair<IConfigurationLinkItem, int[]> linkItem, IList<string> fields);

        /// <summary>
        /// Closes the adapter and releases all resources
        /// </summary>
        void Close();

        /// <summary>
        /// Creates a new <see cref="IWorkItem"/> object in this adapter.
        /// </summary>
        /// <param name="configuration">Configuration of the work item to create</param>
        /// <returns>The new <see cref="IWorkItem"/> object or null if the adapter failed to create work item.</returns>
        /// <exception cref="ConfigurationException">Thrown when the adapter does not support the work item type</exception>
        IWorkItem CreateNewWorkItem(IConfigurationItem configuration);

        /// <summary>
        /// Saves all changes made to work items to its data source
        /// </summary>
        /// <returns>Returns a list of <see cref="ISaveError"/> if errors occurred. The list is empty if no error occurred</returns>
        IList<ISaveError> Save();

        /// <summary>
        /// Validates all work items of the adapter to meet designated rules
        /// </summary>
        /// <returns>Returns a list of <see cref="ISaveError"/> if errors occurred. The list is empty if no error occurred</returns>
        IList<ISaveError> ValidateWorkItems();

        /// <summary>
        /// Gets a web access url for the work item editor
        /// </summary>
        /// TODO Maybe move this to a less general interface?
        Uri GetWorkItemEditorUrl(int workItemId);

        /// <summary>
        /// The method executes all operations.
        /// </summary>
        /// <param name="operations">List of operations to execute.</param>
        void ProcessOperations(IList<IConfigurationTestOperation> operations);

        /// <summary>
        /// Gets the list of available work items from tfs.
        /// </summary>
        ICollection<WorkItem> AvailableWorkItemsFromTFS { get; }
    }
}
