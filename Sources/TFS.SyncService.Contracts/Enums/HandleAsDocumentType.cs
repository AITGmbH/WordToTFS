namespace AIT.TFS.SyncService.Contracts.Enums
{
    /// <summary>
    /// Defines mode for HandleAsDocument processing.
    /// </summary>
    public enum HandleAsDocumentType
    {
        /// <summary>
        /// Process all field values as micro documents, no matter whether contains OLE objects or not.
        /// </summary>
        All,

        /// <summary>
        /// Process only field values as micro documents, which contains only OLE objects.
        /// </summary>
        OleOnDemand
    }
}
