using System;
using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// The method implements operations to work with <see cref="IWord2007TestReportAdapter"/>.
    /// </summary>
    public interface ITestReportHelper
    {
        #region Properties

        /// <summary>
        /// Gets the <see cref="ITfsTestAdapter"/>.
        /// Adapter is used to get additional information.
        /// </summary>
        ITfsTestAdapter TfsTestAdapter { get; }

        /// <summary>
        /// Gets the <see cref="IWord2007TestReportAdapter"/>.
        /// Adapter is used to insert the templates to the word document.
        /// </summary>
        IWord2007TestReportAdapter TestReportAdapter { get; }

        /// <summary>
        /// Gets the <see cref="IConfiguration"/>.
        /// Configuration is used to determine the templates which will be used to insert to the word document.
        /// </summary>
        IConfiguration Configuration { get; }

        /// <summary>
        /// Gets cancelation method - the method is periodically called and the execution is canceled if the cancelation method returns true.
        /// </summary>
        Func<bool> Cancellation { get; }

        #endregion Public properties

        #region Insert to word methods

        /// <summary>
        /// The method inserts the text at the cursor position as heading text with defined level.
        /// </summary>
        /// <param name="text">Text to use.</param>
        /// <param name="headingLevel">Level of the heading. First level is 1.</param>
        void InsertHeadingText(string text, int headingLevel);

        /// <summary>
        /// The method inserts the template for test plan and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testPlanDetail"><see cref="ITfsTestPlanDetail"/> - related test plan.</param>
        void InsertTestPlanTemplate(string templateName, ITfsTestPlanDetail testPlanDetail);

        /// <summary>
        /// The method inserts the template for test suite and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testSuiteDetail"><see cref="ITfsTestSuiteDetail"/> - related test suite.</param>
        void InsertTestSuiteTemplate(string templateName, ITfsTestSuiteDetail testSuiteDetail);

        /// <summary>
        /// The method inserts the template for test case and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testCase"><see cref="ITfsTestCaseDetail"/> - related test case.</param>
        void InsertTestCase(string templateName, ITfsTestCaseDetail testCase);

        /// <summary>
        /// The method inserts the template for shared step and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="sharedStep"><see cref="ITfsSharedStepDetail"/> - related shared step.</param>
        void InsertSharedStep(string templateName, ITfsSharedStepDetail sharedStep);

        /// <summary>
        /// The method inserts the template for test result and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testResult"><see cref="ITfsTestResultDetail"/> - related test result.</param>
        void InsertTestResult(string templateName, ITfsTestResultDetail testResult);

        /// <summary>
        /// The method inserts the template for test configuration and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testConfiguration"><see cref="ITfsTestConfigurationDetail"/> - related test configuration.</param>
        void InsertTestConfiguration(string templateName, ITfsTestConfigurationDetail testConfiguration);

        /// <summary>
        /// The method inserts the template as 'header' for block of templates.
        /// </summary>
        /// <param name="templateName">Name of template to insert.</param>
        void InsertHeaderTemplate(string templateName);

        #endregion Public insert to word methods
    }
}
