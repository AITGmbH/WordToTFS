using System;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationTestCondition"/>.
    /// </summary>
    public class ConfigurationTestCondition : IConfigurationTestCondition
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestCondition"/> class
        /// </summary>
        /// <param name="testCondition">Associated configuration from w2t file.</param>
        public ConfigurationTestCondition(ConditionConfiguration testCondition)
        {
            if (testCondition == null)
                throw new ArgumentNullException("testCondition");
            PropertyToEvaluate = testCondition.PropertyToEvaluate;
            Values = new List<string>();
            if (testCondition.Values != null)
            {
                foreach (var value in testCondition.Values)
                    Values.Add(value);
            }
        }

        #endregion Constructors

        #region Implementation of IConfigurationTestTemplate

        /// <summary>
        /// Gets the property to evaluate and gets the value for condition.
        /// </summary>
        public string PropertyToEvaluate { get; private set; }

        /// <summary>
        /// Gets the list of values that should be used to compare with value of evaluated property.
        /// </summary>
        public IList<string> Values { get; private set; }

        #endregion Implementation of IConfigurationTestTemplate
    }
}
