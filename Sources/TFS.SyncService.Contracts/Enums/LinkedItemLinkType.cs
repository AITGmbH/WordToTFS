namespace AIT.TFS.SyncService.Contracts.Enums
{
    /// <summary>
    /// Defines the type of link between two work items.
    /// </summary>
    public enum LinkedItemLinkType
    {
        /// <summary>
        /// Child means, that the link goes from 'Parent' to 'Child' work item.
        /// </summary>
        Child,

        /// <summary>
        /// Parent means, that the link goes from 'Child' to 'Parent' work item.
        /// </summary>
        Parent
    }
}
