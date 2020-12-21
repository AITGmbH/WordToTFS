using System.Collections.Generic;
using System.Drawing;
using System.Dynamic;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using Microsoft.Office.Interop.Word;

namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Interface defines the functionality of an configuration item.
    /// </summary>
    public interface IConfigurationItem
    {
        /// <summary>
        /// The property gets the information whether the configured item supports the linked items.
        /// </summary>
        bool ConfigurationSupportsLinkedItem { get; }

        /// <summary>
        /// The work item type name for the destination of the sync process
        /// </summary>
        string WorkItemType { get; }

        /// <summary>
        /// Defines reference field name, which must be evaluated to become the sub type of work item.
        /// For example, 'Scenario' is sub type of 'Requirement'.
        /// </summary>
        string WorkItemSubtypeField { get; }

        /// <summary>
        /// Defines value of related sub type of work item.
        /// For example, 'Scenario' is sub type of 'Requirement'.
        /// </summary>
        string WorkItemSubtypeValue { get; }

        /// <summary>
        /// The work item type name for the source of the sync process
        /// </summary>
        string WorkItemTypeMapping { get; }

        /// <summary>
        /// Gets the list of pre-operations.
        /// </summary>
        IList<IConfigurationTestOperation> PreOperations { get; }

        /// <summary>
        /// Gets the list of post-operations.
        /// </summary>
        IList<IConfigurationTestOperation> PostOperations { get; }

        /// <summary>
        /// The regular expression to identify requirement tables. See <see cref="ReqTableCellRow"/> and <see cref="ReqTableCellCol"/>.
        /// </summary>
        string ReqTableIdentifierExpression { get; }

        /// <summary>
        /// The row of the cell to identify requirement tables. See <see cref="ReqTableIdentifierExpression"/> and <see cref="ReqTableCellCol"/>.
        /// </summary>
        int ReqTableCellRow { get; }

        /// <summary>
        /// The column of the cell to identify requirement tables. See <see cref="ReqTableIdentifierExpression"/> and <see cref="ReqTableCellRow"/>.
        /// </summary>
        int ReqTableCellCol { get; }

        /// <summary>
        /// The specific <see cref="IConfigurationFieldItem"/> objects which can be accessed by the destination name
        /// </summary>
        IList<IConfigurationFieldItem> FieldConfigurations { get; }

        /// <summary>
        /// The property gets all configured fields of type <see cref="IConfigurationFieldToLinkedItem"/> to define linked work items.
        /// </summary>
        IList<IConfigurationFieldToLinkedItem> FieldToLinkedItemConfiguration { get; }

        /// <summary>
        /// Gets all configured links.
        /// </summary>
        /// <value>All configured links.</value>
        IList<IConfigurationLinkItem> Links { get; }

        /// <summary>
        /// Image that represents the work item in the UI
        /// </summary>
        Image ImageFile { get; }

        /// <summary>
        /// Related schema of the work item
        /// </summary>
        string RelatedSchema { get; }

        /// <summary>
        /// Related template of the work item. This is used to identify the TFS work item type of the <see cref="IConfigurationItem"/>
        /// abstraction
        /// </summary>
        string RelatedTemplate { get; }

        /// <summary>
        /// Gets template file.
        /// </summary>
        string RelatedTemplateFile { get; }

        /// <summary>
        /// Determines if the Element should be visible in Word
        /// </summary>
        bool HideElementInWord{ get; }

        /// <summary>
        /// This methods can be used to obtain a <see cref="IConverter"/> for a specific <see cref="IField"/>.
        /// The field is identified by the field reference name
        /// The methods uses a dictionary to retrieve the field. Therefore the Access speed is O(1)
        /// </summary>
        /// <param name="fieldName">The <see cref="IField"/> which has to be converted</param>
        /// <returns>Returns the converter for a specific <see cref="IField"/> object</returns>
        IConverter GetConverter(string fieldName);

        /// <summary>
        /// Sets the converter for a given field. This can be used to create context sensitive converters
        /// </summary>
        /// <param name="fieldName">The target field of the conversion</param>
        /// <param name="converter">The converter to be used</param>
        void SetConverter(string fieldName, IConverter converter);

        /// <summary>
        /// Checks if configured transitions should be applied: If word item has no manually changed
        /// state field, it will use transitions rules in which case either word item (if mapped) or
        /// TFS item is changed.
        /// </summary>
        /// <param name="workItemTfs">mapped work item in TFS</param>
        /// <param name="workItemWord">mapped work item in word</param>
        void DoTransition(IWorkItem workItemTfs, IWorkItem workItemWord);
        
        /// <summary>
        /// Level of Header configuration
        /// </summary>
        int Level { get; }






        /// <summary>
        /// Creates a flat clone of the configuration. You can alter the field configuration
        /// of the clone without changing the original configuration.
        /// </summary>
        /// <returns>A flat copy of the configuration</returns>
        IConfigurationItem Clone();
    }
}