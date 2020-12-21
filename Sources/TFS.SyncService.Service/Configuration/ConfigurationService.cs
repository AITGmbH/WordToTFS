using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Factory;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// Class implements the functionality of configuration service - <see cref="IConfigurationService"/>.
    /// </summary>
    internal class ConfigurationService : IConfigurationService
    {
        private readonly Dictionary<object, Configuration> _configurations = new Dictionary<object, Configuration>();

        public IConfiguration GetConfiguration(object document)
        {
            Guard.ThrowOnArgumentNull(document, "document");

            if(!_configurations.ContainsKey(document))
            {
                _configurations.Add(document, new Configuration());
            }

            return _configurations[document];
        }
    }
}