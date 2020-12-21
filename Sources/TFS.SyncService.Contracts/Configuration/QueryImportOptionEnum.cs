namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Definition of different ways to select what work items to load from TFS server
    /// </summary>
    public enum QueryImportOption
    {
        /// <summary>
        /// Use save query to select and load work items.
        /// </summary>
        SavedQuery,

        /// <summary>
        /// Give a list of ids of the work items to load
        /// </summary>
        IDs,

        /// <summary>
        /// Give a text and load work items containing that text
        /// </summary>
        TitleContains
    }
}
