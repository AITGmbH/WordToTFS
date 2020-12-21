using System;
using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.Configuration
{
    public interface ISystemVariableService
    {
        IReadOnlyList<SystemVariable> GetSystemVariables();

        bool TryGetValueByName(string name, out string value);
    }

    public class SystemVariable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="T:System.Object" /> class.
        /// </summary>
        public SystemVariable(string name, string value)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Value cannot be null or whitespace.", "name");
            }

            Name = name;
            Value = value;
        }

        public string Name { get; private set; }
        public string Value { get; private set; }
    }
}
