namespace AIT.TFS.SyncService.Contracts.Enums
{
    /// <summary>
    /// Defines the kind of definition of linked work item
    /// </summary>
    public enum WorkItemBindType
    {
        /// <summary>
        /// The definition of linked work item is in nested table.
        /// </summary>
        NestedTable,

        /// <summary>
        /// The definition of linked work item is in numbered list.
        /// </summary>
        NumberedList,
    }
}
