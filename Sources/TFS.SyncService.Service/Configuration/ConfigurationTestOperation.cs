using System;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements the <see cref="IConfigurationTestOperation"/>.
    /// </summary>
    public class ConfigurationTestOperation : IConfigurationTestOperation
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestOperation"/> class.
        /// </summary>
        /// <param name="operation">Configuration of operation.</param>
        public ConfigurationTestOperation(OperationConfiguration operation)
        {
            if (operation == null)
                throw new ArgumentNullException("operation");

            Type = operation.OperationType;
        }

        #endregion Constructors

        #region Implementation of ICofigurationTestOperation

        /// <summary>
        /// Gets the type of operation.
        /// </summary>
        public OperationType Type { get; private set; }
        
        #endregion Implementation of IConfigurationTestOperation
    }
}
