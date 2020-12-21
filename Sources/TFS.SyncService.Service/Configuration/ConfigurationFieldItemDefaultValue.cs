using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// Default value of a single work item field.
    /// </summary>
    internal class ConfigurationFieldItemDefaultValue : IConfigurationFieldItemDefaultValue
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationFieldItemDefaultValue"/> class.
        /// </summary>
        /// <param name="showName">Name to show in GUI to configure the default value.</param>
        /// <param name="defaultValue">Preconfigured default value of a field.</param>
        public ConfigurationFieldItemDefaultValue(string showName, string defaultValue)
        {
            ShowName = showName;
            DefaultValue = defaultValue;
        }

        #region IConfigurationFieldItemDefaultValue Members

        /// <summary>
        /// Gets the name to show in GUI to configure the default value.
        /// </summary>
        public string ShowName { get; private set; }

        /// <summary>
        /// Gets the preconfigured default value of a field.
        /// </summary>
        public string DefaultValue { get; private set; }

        #endregion
    }
}