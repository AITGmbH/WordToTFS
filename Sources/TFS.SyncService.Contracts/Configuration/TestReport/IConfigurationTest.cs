using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for test report area.
    /// </summary>
    public interface IConfigurationTest
    {
        /// <summary>
        /// Gets the configuration for test specification report area.
        /// </summary>
        IConfigurationTestSpecification ConfigurationTestSpecification { get; }

        /// <summary>
        /// Gets the configuration for test result report area.
        /// </summary>
        IConfigurationTestResult ConfigurationTestResult { get; }

        /// <summary>
        /// Gets the list of all available templates.
        /// </summary>
        IList<IConfigurationTestTemplate> Templates { get; }

        /// <summary>
        /// Gets whether the plugin should set the hyperlink base to the attachment folder
        /// </summary>
        bool SetHyperlinkBase { get; }

        /// <summary>
        /// Gets a value indicating whether to show message boxes when hyperlink base must be adjusted.
        /// </summary>
        bool ShowHyperlinkBaseMessageBoxes
        {
            get;
        }

        /// <summary>
        /// Gets a value indicating whether to expand shared steps for this template.
        /// </summary>
        /// <value>
        ///   <c>true</c> if [shared steps]; otherwise, <c>false</c>.
        /// </value>
        bool ExpandSharedSteps { get; }

        /// <summary>
        /// The method determines the file name of template for the given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the template file name for.</param>
        /// <returns>Determined template file name with full path.</returns>
        string GetFileName(string templateName);

        /// <summary>
        /// The method determines the template name defined as header for the given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the header template for.</param>
        /// <returns>Determined header template name.</returns>
        string GetHeaderTemplate(string templateName);

        /// <summary>
        /// The method determines the conditions of template for the given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the conditions for.</param>
        /// <returns>Determined conditions.</returns>
        IList<IConfigurationTestCondition> GetConditions(string templateName);

        /// <summary>
        /// The method determines the replacements of template for the given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the replacements for.</param>
        /// <returns>Determined replacements.</returns>
        IList<IConfigurationTestReplacement> GetReplacements(string templateName);

        /// <summary>
        /// The method determines the pre-operations for given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the operations for.</param>
        /// <returns>Determined operations. <c>null</c> means that no operation is defined.</returns>
        IList<IConfigurationTestOperation> GetPreOperations(string templateName);

        /// <summary>
        /// The method determines the post-operations for given template name.
        /// </summary>
        /// <param name="templateName">Template name to get the operations for.</param>
        /// <returns>Determined operations. <c>null</c> means that no operation is defined.</returns>
        IList<IConfigurationTestOperation> GetPostOperations(string templateName);


        /// <summary>
        /// Get the PreOperations for a Testresult Report
        /// </summary>
        /// <returns></returns>
        IList<IConfigurationTestOperation> GetPreOperationsForTestResult();

        /// <summary>
        /// Get the PostOperations for a Testresult Report
        /// </summary>
        /// <returns></returns>
        IList<IConfigurationTestOperation> GetPostOperationsForTestResult();


        /// <summary>
        /// Get the PreOperations for a Testspecification Report
        /// </summary>
        /// <returns></returns>
        IList<IConfigurationTestOperation> GetPreOperationsForTestSpecification();


        /// <summary>
        /// Get the PostOperations for a TestSpecification Report
        /// </summary>
        /// <returns></returns>
        IList<IConfigurationTestOperation> GetPostOperationsForTestSpecification();


    }
}
