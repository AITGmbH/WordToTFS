using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts
{
    /// <summary>
    /// Service for working with TFS server.
    /// </summary>
    public interface ITfsService
    {
        /// <summary>
        /// Get all queries for current project.
        /// </summary>
        /// <returns>Project query hierarchy or null when failing to load query hierarchy</returns>
        QueryHierarchy GetWorkItemQueries();

        /// <summary>
        /// Get link types for current project.
        /// </summary>
        /// <returns>A list of link types or null when failed to load the list</returns>
        ICollection<ITFSWorkItemLinkType> GetLinkTypes();

        /// <summary>
        /// Get Area path hierarchy tree.
        /// </summary>
        /// <returns></returns>
        IAreaIterationNode GetAreaPathHierarchy();

        /// <summary>
        /// Get Iteration path hierarchy tree.
        /// </summary>
        /// <returns></returns>
        IAreaIterationNode GetIterationPathHierarchy();

        /// <summary>
        /// Get all work item types for current project.
        /// </summary>
        /// <returns></returns>
        WorkItemTypeCollection GetAllWorkItemTypes();

        /// <summary>
        /// Gets the revision web access URI.
        /// </summary>
        /// <param name="workItem">The work item for which to create the link.</param>
        /// <param name="revision">The revision of the work item for which to create link.</param>
        /// <returns></returns>
        Uri GetRevisionWebAccessUri(IWorkItem workItem, int revision);

        /// <summary>
        /// Gets the category collection of the current project.
        /// </summary>
        /// <value>
        /// The category collection.
        /// </value>
        CategoryCollection CategoryCollection { get; }

        /// <summary>
        /// Name of the team project the adapter is connected to
        /// </summary>
        string ProjectName { get; }

        /// <summary>
        /// Name / URL of the TFS server
        /// </summary>
        string ServerName { get; }

        /// <summary>
        /// Gets the list of all configuration field items that are available 
        /// on the server.
        /// </summary>
        FieldDefinitionCollection FieldDefinitions
        {
            get;
        }
    }
}
