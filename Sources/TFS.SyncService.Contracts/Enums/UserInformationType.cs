namespace AIT.TFS.SyncService.Contracts.Enums
{
    /// <summary>
    /// Describes what kind of user information to display
    /// </summary>
    public enum UserInformationType
    {
        /// <summary>
        /// Work item was published successfully
        /// </summary>
        Success,

        /// <summary>
        /// Non-Fatal error, most likely a configuration error.
        /// </summary>
        Warning,

        /// <summary>
        /// Severe error that is not caused by a configuration error.
        /// </summary>
        Error,

        /// <summary>
        /// Work item did not need any changed because it was unmodified.
        /// </summary>
        Unmodified,

        /// <summary>
        /// Work item was not synchronized because of a revision conflict
        /// </summary>
        Conflicting
    }
}