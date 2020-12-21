namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// AttachmentFolderMode allows to specficy if attachments should be downloaded to folders that contain Guid or TestSuiteId or nothing (Mode: WithGuid/WithoutGuid/BasedOnTestSuite)
    /// </summary>
    public enum AttachmentFolderMode
    {
        /// <summary>
        /// Name of attachment contains Guid.
        /// </summary>
        WithGuid,

        /// <summary>
        /// Name of attachment does not contain Guid.
        /// </summary>
        WithoutGuid,

        /// <summary>
        /// Name of attachment is based on the test suite.
        /// </summary>
        BasedOnTestSuite
    }
}
