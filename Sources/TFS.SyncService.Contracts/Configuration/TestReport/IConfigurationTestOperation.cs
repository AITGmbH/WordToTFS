namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for one operation before / after insert of template.
    /// </summary>
    public interface IConfigurationTestOperation
    {
        /// <summary>
        /// Gets the type of operation.
        /// </summary>
        OperationType Type { get; }
    }
}
