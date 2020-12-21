namespace AIT.TFS.SyncService.Contracts.Enums
{
    /// <summary>
    /// Defines the direction of transfer for the work items.
    /// </summary>
    public enum Direction
    {
        /// <summary>
        /// Synchronize from word to server
        /// </summary>
        OtherToTfs,

        /// <summary>
        /// Synchronize from server to word.
        /// </summary>
        TfsToOther,

        /// <summary>
        /// Special direction, used to set a field in new work item created on server.
        /// </summary>
        SetInNewTfsWorkItem,

        /// <summary>
        /// Special direction, used to set a field in new work item created on word (so its for newly imported items).
        /// </summary>
        GetOnly,

        /// <summary>
        /// Field is published only, but never refreshed from server
        /// </summary>
        PublishOnly
    }
}