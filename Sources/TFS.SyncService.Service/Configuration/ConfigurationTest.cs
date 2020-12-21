using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationTest"/> - configuration for test report area.
    /// </summary>
    public class ConfigurationTest : IConfigurationTest
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTest"/> class.
        /// </summary>
        public ConfigurationTest()
        {
            Folder = string.Empty;
            ConfigurationTestSpecification = new ConfigurationTestSpecification();
            ConfigurationTestResult = new ConfigurationTestResult();
            Templates = new List<IConfigurationTestTemplate>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTest"/> class.
        /// </summary>
        /// <param name="folder">Folder containing all template files.</param>
        /// <param name="testConfiguration">Associated test area configuration from w2t file.</param>
        public ConfigurationTest(string folder, TestConfiguration testConfiguration) : this()
        {
            Folder = folder;
            TestConfiguration = testConfiguration;
            if (testConfiguration != null)
            {
                // TODO ShowHyperlinkBaseMessageBoxes should be configurable in the xml file.
                ShowHyperlinkBaseMessageBoxes = true;
                SetHyperlinkBase = testConfiguration.SetHyperlinkBase;
                ExpandSharedSteps = testConfiguration.ExpandSharedSteps;
                ConfigurationTestSpecification = new ConfigurationTestSpecification(TestConfiguration.TestSpecificationConfiguration);
                ConfigurationTestResult = new ConfigurationTestResult(TestConfiguration.TestResultConfiguration);
                Templates = new List<IConfigurationTestTemplate>();
                foreach (var template in TestConfiguration.TemplatesConfiguration) Templates.Add(new ConfigurationTestTemplate(template));
            }
        }

        #endregion Constructors

        #region Internal properties

        /// <summary>
        /// Gets the test area configuration read from w2t file. Set only in one constructor.
        /// </summary>
        internal TestConfiguration TestConfiguration { get; private set; }

        /// <summary>
        /// Gets the folder containing all template files.
        /// </summary>
        internal string Folder { get; private set; }

        #endregion Internal properties

        #region Implementation of IConfigurationTest

        /// <summary>
        /// Gets the configuration for test specification report area.
        /// </summary>
        public IConfigurationTestSpecification ConfigurationTestSpecification { get; private set; }

        /// <summary>
        /// Gets the configuration for test result report area.
        /// </summary>
        public IConfigurationTestResult ConfigurationTestResult { get; private set; }

        /// <summary>
        /// Gets the list of all available templates.
        /// </summary>
        public IList<IConfigurationTestTemplate> Templates { get; private set; }

        /// <summary>
        /// Gets whether the plugin should set the hyperlink base
        ///  to the attachment folder
        /// </summary>
        public bool SetHyperlinkBase { get; private set; }

        /// <summary>
        /// Gets a value indicating whether to show message boxes when hyperlink base must be adjusted.
        /// </summary>
        public bool ShowHyperlinkBaseMessageBoxes
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets a value indicating whether to expand shared steps for this template.
        /// </summary>
        /// <value>
        ///   <c>true</c> if [shared steps]; otherwise, <c>false</c>.
        /// </value>
        public bool ExpandSharedSteps { get; private set; }

        /// <summary>
        /// The method determines the file name of template for the given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the template file name for.</param>
        /// <returns>Determined template file name with full path.</returns>
        public string GetFileName(string templateName)
        {
            var template = (from t in Templates where t.Name == templateName select t).FirstOrDefault();
            if (template == null)
                return null;
            return Path.Combine(Folder, template.FileName);
        }

        /// <summary>
        /// The method determines the template name defined as header for the given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the header template for.</param>
        /// <returns>Determined header template name.</returns>
        public string GetHeaderTemplate(string templateName)
        {
            var template = (from t in Templates where t.Name == templateName select t).FirstOrDefault();
            if (template == null)
                return null;
            return template.HeaderTemplate;
        }

        /// <summary>
        /// The method determines the conditions of template for the given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the conditions for.</param>
        /// <returns>Determined conditions.</returns>
        public IList<IConfigurationTestCondition> GetConditions(string templateName)
        {
            var template = (from t in Templates where t.Name == templateName select t).FirstOrDefault();
            if (template == null)
                throw new ConfigurationException($"There template {templateName} was not found in the configuration file");
            return template.Conditions;
        }

        /// <summary>
        /// The method determines the replacements of template for the given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the replacements for.</param>
        /// <returns>Determined replacements.</returns>
        public IList<IConfigurationTestReplacement> GetReplacements(string templateName)
        {
            var template = (from t in Templates where t.Name == templateName select t).FirstOrDefault();
            if (template == null)
                return null;
            return template.Replacements;
        }

        /// <summary>
        /// The method determines the pre-operations for given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the operations for.</param>
        /// <returns>Determined operations. <c>null</c> means that no operation is defined.</returns>
        public IList<IConfigurationTestOperation> GetPreOperations(string templateName)
        {
            var template = (from t in Templates where t.Name == templateName select t).FirstOrDefault();
            if (template == null)
                return null;
            return template.PreOperations;
        }



        /// <summary>
        /// The method determines the post-operations for given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the operations for.</param>
        /// <returns>Determined operations. <c>null</c> means that no operation is defined.</returns>
        public IList<IConfigurationTestOperation> GetPostOperations(string templateName)
        {
            var template = (from t in Templates where t.Name == templateName select t).FirstOrDefault();
            if (template == null)
                return null;
            return template.PostOperations;
        }

        /// <summary>
        /// Get the pre-operations for the testresult report
        /// </summary>
        /// <returns></returns>
        public IList<IConfigurationTestOperation> GetPreOperationsForTestResult()
        {
            return ConfigurationTestResult.PreOperations;
        }

        /// <summary>
        /// Get the post-operations for the testresult report
        /// </summary>
        /// <returns></returns>
        public IList<IConfigurationTestOperation> GetPostOperationsForTestResult()
        {
            return ConfigurationTestResult.PostOperations;
        }

        /// <summary>
        /// Get the pre-operations for the testspecification report
        /// </summary>
        /// <returns></returns>
        public IList<IConfigurationTestOperation> GetPreOperationsForTestSpecification()
        {
            return ConfigurationTestSpecification.PreOperations;
        }

        /// <summary>
        /// Get the post-operations for the testspecification report
        /// </summary>
        /// <returns></returns>
        public IList<IConfigurationTestOperation> GetPostOperationsForTestSpecification()
        {
            return ConfigurationTestSpecification.PostOperations;
        }

        #endregion Implementation of IConfigurationTest
    }
}
