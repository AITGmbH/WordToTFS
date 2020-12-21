namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Interface that describes a variable from the config
    /// </summary>
    public interface IVariable
    {
        /// <summary>
        /// Gets the Name of the variable
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// Gets the value of the variable
        /// </summary>
        string Value { get; set; }
    }
}
