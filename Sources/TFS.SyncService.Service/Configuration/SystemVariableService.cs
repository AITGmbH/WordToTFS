using System.Collections.Generic;

using AIT.TFS.SyncService.Common;
using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Service.Configuration
{
    using System;

    public class SystemVariableService : ISystemVariableService
    {
        private static readonly List<SystemVariable> SystemVariables;

        static SystemVariableService()
        {
            SystemVariables = new List<SystemVariable>();

            var version = ProductInformation.AssemblyVersion;

            SystemVariables.Add(new SystemVariable("WordToTFS.Version", version));
        }

        public IReadOnlyList<SystemVariable> GetSystemVariables()
        {
            return SystemVariables;
        }

        public bool TryGetValueByName(string name, out string value)
        {
            foreach (var systemVariable in SystemVariables)
            {
                if (string.Equals(systemVariable.Name, name, StringComparison.Ordinal))
                {
                    value = systemVariable.Value;
                    return true;
                }
            }

            value = null;
            return false;
        }
    }
}
