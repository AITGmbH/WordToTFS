using System;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Interface defines functionality for one field in work item.
    /// </summary>
    public interface IConfigurationLinkItem
    {
        /// <summary>
        /// Gets the name of the link type.
        /// </summary>
        string LinkValueType { get; }

        /// <summary>
        /// Determines the copy direction of the field
        /// </summary>
        Direction Direction { get; }

        /// <summary>
        /// Table column index
        /// </summary>
        int ColIndex { get; }

        /// <summary>
        /// Table column index
        /// </summary>
        int RowIndex { get; }

        /// <summary>
        /// Gets a value indicating whether to replace all existing links.
        /// </summary>
        bool Overwrite { get; }

        /// <summary>
        /// A character or string used to separate multiple links of the same type, e.g. ","
        /// </summary>
        string LinkSeparator { get; }

        /// <summary>
        /// A composite format string to be used to format a single link. The format parameters
        /// </summary>
        string LinkFormat { get; }

        /// <summary>
        /// Gets the type of work item to which to link automatically when publishing
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "LinkWork")]
        string AutomaticLinkWorkItemType
        {
            get;
        }

        /// <summary>
        /// Gets an optional field used to select the automatically linked work item
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "LinkWork")]
        string AutomaticLinkWorkItemSubtypeField
        {
            get;
        }

        /// <summary>
        /// Gets the value the optional field has to have
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "LinkWork")]
        string AutomaticLinkWorkItemSubtypeValue
        {
            get;
        }

        /// <summary>
        /// Gets whether to suppress warnings when no automatic link target was found
        /// </summary>
        bool AutomaticLinkSuppressWarnings
        {
            get;
        }

        /// <summary>
        /// Formats a work item by the set <see cref="LinkFormat"/>
        /// </summary>
        /// <param name="workItem">Work item to format.</param>
        /// <returns>string representing the work item.</returns>
        string Format(IWorkItem workItem);

        /// <summary>
        /// returns all fields that are necessary to format links
        /// </summary>
        IEnumerable<string> GetLinkFormatRequiredFields();

        /// <summary>
        /// Extracts the work item id from a formatted string. Uses the set <see cref="LinkFormat"/> to find id.
        /// </summary>
        /// <param name="formattedRepresentation">string representing the work item or only an id.</param>
        /// <returns>work item id or Null if the string is not correctly formatted and not an id.</returns>
        /// <exception cref="ConfigurationException">If the id could not be extracted from the formatted string</exception>
        /// <exception cref="OverflowException">When the value in the id field could not be casted into an int</exception>
        int? GetWorkItemId(string formattedRepresentation);

        IEnumerable<string> GetLinkedWorkItemTypes();

        string LinkedWorkItemTypes
        {
            get;
        } 
    }
}