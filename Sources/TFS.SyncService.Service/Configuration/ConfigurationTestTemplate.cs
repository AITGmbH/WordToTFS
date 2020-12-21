using System;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// Implementation of <see cref="IConfigurationTestTemplate"/>
    /// </summary>
    public class ConfigurationTestTemplate : IConfigurationTestTemplate
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestTemplate"/> class.
        /// </summary>
        public ConfigurationTestTemplate()
        {
            Conditions = new List<IConfigurationTestCondition>();
            Replacements = new List<IConfigurationTestReplacement>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestTemplate"/> class.
        /// </summary>
        /// <param name="template">Associated configuration from w2t file.</param>
        public ConfigurationTestTemplate(TemplateConfiguration template)
        {
            if (template == null)
                throw new ArgumentNullException("template");
            Name = template.Name;
            HeaderTemplate = template.HeaderTemplate;
            FileName = template.FileName;
            Conditions = new List<IConfigurationTestCondition>();
            Replacements = new List<IConfigurationTestReplacement>();
            foreach (var condition in template.Conditions)
                Conditions.Add(new ConfigurationTestCondition(condition));
            foreach (var replace in template.Replacements)
                Replacements.Add(new ConfigurationTestReplacement(replace));
            if (template.PreOperations != null && template.PreOperations.Count > 0)
            {
                PreOperations = new List<IConfigurationTestOperation>();
                foreach (var operation in template.PreOperations)
                    PreOperations.Add(new ConfigurationTestOperation(operation));
            }
            if (template.PostOperations != null && template.PostOperations.Count > 0)
            {
                PostOperations = new List<IConfigurationTestOperation>();
                foreach (var operation in template.PostOperations)
                    PostOperations.Add(new ConfigurationTestOperation(operation));
            }
        }

        #endregion Constructors

        #region Implementation of IConfigurationTestTemplate

        /// <summary>
        /// Gets the name of template.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the template name for header of block.
        /// </summary>
        public string HeaderTemplate { get; private set; }

        /// <summary>
        /// Gets the file name of template.
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        /// Gets the list of pre-operations.
        /// </summary>
        public IList<IConfigurationTestOperation> PreOperations { get; private set; }

        /// <summary>
        /// Gets the list of post-operations.
        /// </summary>
        public IList<IConfigurationTestOperation> PostOperations { get; private set; }

        /// <summary>
        /// Gets the list of conditions in one template.
        /// </summary>
        public IList<IConfigurationTestCondition> Conditions { get; private set; }

        /// <summary>
        /// Gets the list of replacements in one template.
        /// </summary>
        public IList<IConfigurationTestReplacement> Replacements { get; private set; }

        #endregion Implementation of IConfigurationTestTemplate
    }
}
