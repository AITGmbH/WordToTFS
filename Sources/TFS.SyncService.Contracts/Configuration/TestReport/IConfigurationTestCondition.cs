using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for one condition in template.
    /// </summary>
    public interface IConfigurationTestCondition
    {
        /// <summary>
        /// Gets the property to evaluate and gets the value for condition.
        /// </summary>
        string PropertyToEvaluate { get; }

        /// <summary>
        /// Gets the list of values that should be used to compare with value of evaluated property.
        /// </summary>
        IList<string> Values { get; }
    }
}
