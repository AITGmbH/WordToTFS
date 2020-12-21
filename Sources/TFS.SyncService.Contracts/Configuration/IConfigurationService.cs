using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Interface defines the functionality of configuration service.
    /// </summary>
    public interface IConfigurationService
    {
        /// <summary>
        /// Returns the configuration for a given document.
        /// </summary>
        IConfiguration GetConfiguration(object document);
    }
}