namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Interface defines functionality for default value object of one field in work item.
    /// </summary>
    public interface IConfigurationFieldItemDefaultValue
    {
        /// <summary>
        /// Gets the name to show in GUI to configure the default value.
        /// </summary>
        string ShowName { get; }

        /// <summary>
        /// Gets the preconfigured default value of a field.
        /// </summary>
        string DefaultValue { get; }
    }
}