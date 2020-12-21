using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Service.Configuration.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationFieldToLinkedItem"/>
    /// and represents the configuration for all linked work items.
    /// </summary>
    public class ConfigurationFieldToLinkedItem : IConfigurationFieldToLinkedItem
    {
        #region Constructors
        /// <summary>
        /// Constructor creates an instance of the class <see cref="ConfigurationFieldToLinkedItem"/>
        /// </summary>
        /// <param name="copyFrom">Object of type <see cref="MappingFieldToLinkedItem"/> to copy the values for properties from.</param>
        public ConfigurationFieldToLinkedItem(MappingFieldToLinkedItem copyFrom)
        {
            if (copyFrom == null)
                return;

            LinkedWorkItemType = copyFrom.LinkedWorkItemType;
            LinkType = copyFrom.LinkType;
            WorkItemBindType = copyFrom.WorkItemBindType;
            ColIndex = copyFrom.TableCol;
            RowIndex = copyFrom.TableRow;
            IList<IConfigurationFieldAssignment> fields = new List<IConfigurationFieldAssignment>();
            if (copyFrom.FieldAssignments != null)
            {
                foreach (var field in copyFrom.FieldAssignments)
                {
                    if (field == null)
                        continue;
                    fields.Add(new ConfigurationFieldAssignment(field));
                }
            }
            FieldAssignmentConfiguration = fields;
        }
        #endregion Constructors

        #region Implementation of IConfigurationFieldToLinkedItem
        /// <summary>
        /// The property gets the name of work item that should be used as linked work item.
        /// </summary>
        public string LinkedWorkItemType { get; private set; }

        /// <summary>
        /// The property gets the type of the relationship between original work item and linked work item.
        /// </summary>
        /// <value>
        /// <c>Parent</c> - the linked work item is parent to the work item where is this configuration defined.
        /// <c>Child</c> - the linked work item is child to the work item where is this configuration defined.
        /// </value>
        public LinkedItemLinkType LinkType { get; private set; }

        /// <summary>
        /// The property gets the kind of definition for the linked work item.
        /// </summary>
        public WorkItemBindType WorkItemBindType { get; private set; }

        /// <summary>
        /// The property gets all configured fields of type <see cref="IConfigurationFieldToLinkedItem"/> to define linked work items.
        /// </summary>
        public IList<IConfigurationFieldAssignment> FieldAssignmentConfiguration { get; private set; }

        /// <summary>
        /// The property gets the index of column in the table where is the definition for the linked work items.
        /// </summary>
        public int ColIndex { get; private set; }

        /// <summary>
        /// The property gets the index of row in the table where is the definition for the linked work items.
        /// </summary>
        public int RowIndex { get; private set; }
        #endregion Implementation of IConfigurationFieldToLinkedItem
    }
}
