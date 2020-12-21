namespace AIT.TFS.SyncService.Contracts.Enums
{
    /// <summary>
    /// Possible synchronization states between a work item in Word and the same work item on TFS
    /// </summary>
    public enum SynchronizationState
    {
        /// <summary>
        /// SynchronizationState is unknown or being queried
        /// </summary>
        Unknown,

        /// <summary>
        /// Item is not imported
        /// </summary>
        NotImported,

        /// <summary>
        /// Item exists in Word but it outdated
        /// </summary>
        Outdated,

        /// <summary>
        /// Item exists in Word but some of its fields have values different than the TFS version. Because neither are the fields publishable, nor is the value changed in TFS this situation will remain until refreshed.
        /// </summary>
        Differing,
            
        /// <summary>
        /// Item exists in Word but both the Word and the TFS version where changed and cannot be merged without conflict.
        /// </summary>
        DivergedWithConflicts,

        /// <summary>
        /// Item exists in Word but both the Word and the TFS version where changed, but can be merged without conflict.
        /// </summary>
        DivergedWithoutConflicts,

        /// <summary>
        /// Item exists and is changed, but the changes are not reflected in TFS
        /// </summary>
        Dirty,

        /// <summary>
        /// Item exists and is up to date
        /// </summary>
        UpToDate,

        /// <summary>
        /// Item is new in Word and has not yet been synchronized to TFS
        /// </summary>
        New,

        /// <summary>
        /// Synchronization state query operation was aborted
        /// </summary>
        Aborted
    }
}
