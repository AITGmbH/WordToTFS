#region Usings
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
using AIT.TFS.SyncService.Common.Helper;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Properties;
using AIT.TFS.SyncService.Service.InfoStorage;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.Office.Interop.Word;
#endregion

namespace AIT.TFS.SyncService.Model.TestReport
{
    /// <summary>
    /// The method implements operations to work with <see cref="IWord2007TestReportAdapter"/>.
    /// </summary>
    /// // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
    public class TestReportHelper : ITestReportHelper
    {
        #region Private fields
        private readonly ISystemVariableService _systemVariableService;
        private readonly Dictionary<int, string> _resolutionStates;
        private ITfsTestSuiteDetail _testSuite;
        private ITfsTestCaseDetail _testCase;
        private IList<string> _attachmentNames = new List<string>();
        private int _numberOfStepAttachmentsPerTestCase;
        private IDictionary<int, List<string>> _stepsAndAttachmentsPerTestCase = new Dictionary<int, List<string>>();
        private const string TEST_CASE_PREFIX = "TC_";
        private const string STEP_PREFIX = "_Step_";
        private static bool _isDuplicateFolderCreated;
        private static string _guid;
        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TestReportHelper"/> class.
        /// </summary>
        /// <param name="tfsTestAdapter">Adapter will be used to get additional information.</param>
        /// <param name="testReportAdapter">Adapter will be used to insert the templates to the word document.</param>
        /// <param name="configuration">Configuration will be used to obtain the templates which will be used to insert to the word document.</param>
        /// <param name="cancellation">cancellation method - the method is periodically called and the execution is canceled if the cancellation method returns true.</param>
        public TestReportHelper(ITfsTestAdapter tfsTestAdapter, IWord2007TestReportAdapter testReportAdapter, IConfiguration configuration, Func<bool> cancellation)
        {
            if (tfsTestAdapter == null)
                throw new ArgumentNullException("tfsTestAdapter");
            if (testReportAdapter == null)
                throw new ArgumentNullException("testReportAdapter");
            if (configuration == null)
                throw new ArgumentNullException("configuration");

            _systemVariableService = SyncServiceFactory.GetService<ISystemVariableService>();


            TfsTestAdapter = tfsTestAdapter;
            TestReportAdapter = testReportAdapter;
            Configuration = configuration;
            Cancellation = cancellation;
            _resolutionStates = new Dictionary<int, string>();
            if (TfsTestAdapter.TestManagement != null && TfsTestAdapter.TestManagement.TestResolutionStates != null)
            {
                _resolutionStates = TfsTestAdapter.TestManagement.TestResolutionStates.Query().ToDictionary(x => x.Id, x => x.Name);
            }
            // "None" resolution state is not on the list of the resolution states returned from tfs,
            // so we have to add it manually
            _resolutionStates.Add(0, Resources.NoneResolutionState);

        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the <see cref="ITfsTestAdapter"/>.
        /// Adapter is used to get additional information.
        /// </summary>
        public ITfsTestAdapter TfsTestAdapter { get; private set; }

        /// <summary>
        /// Gets the <see cref="IWord2007TestReportAdapter"/>.
        /// Adapter is used to insert the templates to the word document.
        /// </summary>
        public IWord2007TestReportAdapter TestReportAdapter { get; private set; }

        /// <summary>
        /// Gets the <see cref="IConfiguration"/>.
        /// Configuration is used to determine the templates which will be used to insert to the word document.
        /// </summary>
        public IConfiguration Configuration { get; private set; }

        /// <summary>
        /// Gets cancellation method - the method is periodically called and the execution is canceled if the cancellation method returns true.
        /// </summary>
        public Func<bool> Cancellation { get; private set; }

        /// <summary>
        /// This property is set by the model. This is necessary to limit the results to the options of the UI
        /// </summary>
        public bool IncludeMostRecentTestResults
        {
            private get;
            set;
        }

        /// <summary>
        /// This property is set by the model. This is necessary to limit the results to the options of the UI
        /// </summary>
        public bool IncludeOnlyMostRecentTestResultForAllConfigurations
        {
            private get;
            set;
        }

        public ITfsTestConfiguration CurrentTestConfiguration
        {
            private get;
            set;
        }


        #endregion Public properties

        #region Public insert to word methods

        /// <summary>
        /// The method inserts the text at the cursor position as heading text with defined level.
        /// </summary>
        /// <param name="text">Text to use.</param>
        /// <param name="headingLevel">Level of the heading. First level is 1.</param>
        public void InsertHeadingText(string text, int headingLevel)
        {
            TestReportAdapter.InsertHeadingText(text, headingLevel);
        }

        /// <summary>
        /// The method inserts the template for test plan and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testPlanDetail"><see cref="ITfsTestPlanDetail"/> - related test plan.</param>
        public void InsertTestPlanTemplate(string templateName, ITfsTestPlanDetail testPlanDetail)
        {
            Insert(templateName, testPlanDetail);
        }

        /// <summary>
        /// The method inserts the template for test suite and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testSuiteDetail"><see cref="ITfsTestSuiteDetail"/> - related test suite.</param>
        public void InsertTestSuiteTemplate(string templateName, ITfsTestSuiteDetail testSuiteDetail)
        {
            _testSuite = testSuiteDetail;
            _attachmentNames.Clear();
            Insert(templateName, testSuiteDetail);
        }

        /// <summary>
        /// The method inserts the template for test case and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testCase"><see cref="ITfsTestCaseDetail"/> - related test case.</param>
        public void InsertTestCase(string templateName, ITfsTestCaseDetail testCase)
        {
            _testCase = testCase;
            _numberOfStepAttachmentsPerTestCase = 0;
            _stepsAndAttachmentsPerTestCase.Clear();
            Insert(templateName, testCase);
        }

        /// <summary>
        /// The method inserts the template for shared steps and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of the template.</param>
        /// <param name="sharedStep">The shared step.</param>
        public void InsertSharedStep(string templateName, ITfsSharedStepDetail sharedStep)
        {
            Insert(templateName, sharedStep);
        }

        /// <summary>
        /// The method inserts the template for test result and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testResult"><see cref="ITfsTestResultDetail"/> - related test result.</param>
        public void InsertTestResult(string templateName, ITfsTestResultDetail testResult)
        {
            Insert(templateName, testResult);
        }

        /// <summary>
        /// The method inserts the template for test configuration and replaces all bookmarks.
        /// </summary>
        /// <param name="templateName">Name of template to use for insert and replacement.</param>
        /// <param name="testConfiguration"><see cref="ITfsTestConfigurationDetail"/> - related test configuration.</param>
        public void InsertTestConfiguration(string templateName, ITfsTestConfigurationDetail testConfiguration)
        {
            Insert(templateName, testConfiguration);
        }

        /// <summary>
        /// The method inserts the template as 'header' for block of templates.
        /// </summary>
        /// <param name="templateName">Name of template to insert.</param>
        public void InsertHeaderTemplate(string templateName)
        {
            if (string.IsNullOrEmpty(templateName))
                return;
            TestReportAdapter.InsertFile(
                 Configuration.ConfigurationTest.GetFileName(templateName),
                 Configuration.ConfigurationTest.GetPreOperations(templateName),
                 Configuration.ConfigurationTest.GetPostOperations(templateName));
        }

        /// <summary>
        /// The method inserts a given summary template at the end of the document
        /// </summary>
        /// <param name="templateName"></param>
        /// <param name="boundElement"></param>
        private void InsertSummaryPageTemplate(string templateName, ITfsPropertyValueProvider boundElement)
        {
            if (string.IsNullOrEmpty(templateName))
                return;
            TestReportAdapter.InsertFile(
                Configuration.ConfigurationTest.GetFileName(templateName),
                Configuration.ConfigurationTest.GetPreOperations(templateName),
                Configuration.ConfigurationTest.GetPostOperations(templateName));

            var expandSharedSteps = Configuration.ConfigurationTest.ExpandSharedSteps;


            //if test plan detail is available, replace the bookmarks
            if (boundElement != null)
            {

                // Store cursor position (it is at the end of inserted template)
                TestReportAdapter.PushCursorPosition();

                // Replace bookmarks
                var replacements = Configuration.ConfigurationTest.GetReplacements(templateName);
                Replace(TestReportAdapter, replacements, boundElement, expandSharedSteps);

                // Restore cursor position (it is at the position of last replaced bookmark)
                TestReportAdapter.PopCursorPosition();

                // Remove not stored bookmarks
                //TestReportAdapter.RemoveBookmarks();
            }
        }

        #endregion Public insert to word methods

        #region Private methods

        private IEnumerable EnumerableExpander(IEnumerable enumerable)
        {
            return TfsTestAdapter.ExpandEnumerable(enumerable);
        }

        #endregion Private methods

        #region Private replace method

        /// <summary>
        /// Inserts a template-file and replaces its bookmarks with values of the bound element
        /// </summary>
        private void Insert(string templateName, ITfsPropertyValueProvider boundElement)
        {
            if (boundElement == null) throw new ArgumentNullException("boundElement");

            // templateName probably intentionally left blank in the configuration
            if (string.IsNullOrEmpty(templateName))
            {
                return;
            }

            // Store bookmarks
            TestReportAdapter.StoreOriginalBookmarks();

            // Insert template
            TestReportAdapter.InsertFile(
                Configuration.ConfigurationTest.GetFileName(templateName),
                Configuration.ConfigurationTest.GetPreOperations(templateName),
                Configuration.ConfigurationTest.GetPostOperations(templateName));

            // Store cursor position (it is at the end of inserted template)
            TestReportAdapter.PushCursorPosition();

            // Replace bookmarks
            var expandSharedSteps = Configuration.ConfigurationTest.ExpandSharedSteps;
            var replacements = Configuration.ConfigurationTest.GetReplacements(templateName);
            Replace(TestReportAdapter, replacements, boundElement, expandSharedSteps);

            // Restore cursor position (it is at the position of last replaced bookmark)
            TestReportAdapter.PopCursorPosition();

            // Remove not stored bookmarks
            TestReportAdapter.RemoveBookmarks();
        }


        /// <summary>
        /// Process operations 
        /// </summary>
        /// <param name="operations"></param>
        public void ProcessOperations(IList<IConfigurationTestOperation> operations)
        {

            //Process Operation
            TestReportAdapter.ProcessOperations(operations);
            //store cursor position
            TestReportAdapter.PushCursorPosition();

        }

        /// <summary>
        /// Create a summary page
        /// </summary>
        public void CreateSummaryPage(Document wordDocument, ITfsTestPlan testPlan)
        {
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(wordDocument);
            var summaryPageTemplate = config.ConfigurationTest.ConfigurationTestResult.SummaryPageTemplate;
            if (summaryPageTemplate != null)
            {
                // Get the detailed test plan. If no testplan is provided
                var testPlanDetail = TfsTestAdapter.GetTestPlanDetail(testPlan);

                //Print out the summary page

                InsertSummaryPageTemplate(summaryPageTemplate, testPlanDetail);
            }
        }

        /// <summary>
        ///  Replaces all bookmarks as defined in a configuration
        /// </summary>
        private void Replace(IWord2007TestReportAdapter adapter, IEnumerable<IConfigurationTestReplacement> replacements, ITfsPropertyValueProvider boundElement, bool expandSharedSteps, TfsTestAction testAction = null)
        {
            foreach (var replacement in replacements)
            {
                try
                {
                    switch (replacement.FieldValueType)
                    {
                        case (FieldValueType.BasedOnVariable):
                            var variable = Configuration.Variables.FirstOrDefault(x => x.Name == replacement.VariableName);
                            if (variable == null)
                            {
                                throw new ConfigurationException(string.Format(Resources.ReplaceUnknownVariableName, replacement.VariableName, replacement.Bookmark));
                            }
                            adapter.ReplaceBookmarkText(replacement.Bookmark, variable.Value, PropertyValueFormat.PlainText, string.Empty);
                            continue;
                        case (FieldValueType.BasedOnSystemVariable):
                            string value;
                            if (_systemVariableService.TryGetValueByName(replacement.VariableName, out value) == false)
                            {
                                throw new ConfigurationException(string.Format(Resources.ReplaceUnknownSystemVariableName, replacement.VariableName, replacement.Bookmark));
                            }
                            adapter.ReplaceBookmarkText(replacement.Bookmark, value, PropertyValueFormat.PlainText, string.Empty);
                            continue;
                        case FieldValueType.HTML:
                        case FieldValueType.BasedOnFieldType:
                        case FieldValueType.DropDownList:
                            throw new ConfigurationException(string.Format(Resources.ReplaceFieldValueTypeNotSupported, replacement.FieldValueType));
                    }

                    //Check if the property is a parametrized property
                    if (replacement.PropertyToEvaluate.Contains("LinkedWorkItemsForTestResult"))
                    {
                        //Check if paramters are provided with the property
                        if (!(string.IsNullOrEmpty(replacement.Parameters)))
                        {
                            //Search the arguments
                            var matches = Regex.Matches(replacement.Parameters, @"\{(.*?)\}");
                            if (matches != null && matches.Count != 0)
                            {
                                //Replace the property of the current object with list of workItems
                                var value = GetValuesForParametrizedQuery(boundElement, matches);
                                (boundElement as ITfsTestResultDetail).LinkedWorkItemsForTestResult = value;
                            }
                        }
                        else
                        {
                            throw new ConfigurationException("You are trying to use the property LinkedWorkItemsForTestResults with an appropriate Parameters property. Please refer to the manual for the correct setup of this property");
                        }
                    }

                    if (replacement.PropertyToEvaluate.Contains("TestParametersWithAllValues"))
                    {
                        if (boundElement.GetType() == typeof(TfsTestCaseDetail))
                        {
                            var value = GetParametersForTestCase(boundElement);
                            (boundElement as ITfsTestCaseDetail).TestParametersWithAllValues = value;
                        }
                        else
                        {
                            throw new ConfigurationException("You are trying to use the property TestParametersWithAllValues for an object that does not contain this property.");
                        }
                    }

                    //Search user defined filter
                    foreach (var objectQuery in Configuration.ObjectQueries)
                    {
                        if (objectQuery.Name.Contains(replacement.PropertyToEvaluate))
                        {
                            if ((objectQuery.DestinationElements.Any(x => x.ItemName.Equals("Builds"))))
                            {
                                var value = GetBuildsForTestManagmentObject(boundElement, objectQuery, expandSharedSteps);
                                boundElement.AddCustomObjects(replacement.PropertyToEvaluate, value);
                            }
                            if ((objectQuery.DestinationElements.Any(x => x.ItemName.Equals("Test Cases"))))
                            {
                                var value = GetLinkedTestCasesForTestManagmentObject(boundElement, objectQuery, expandSharedSteps);
                                boundElement.AddCustomObjects(replacement.PropertyToEvaluate, value);
                            }
                            else if (objectQuery.DestinationElements.Any(x => x.ItemName.Equals("Test Results")))
                            {
                                var value = GetLinkedTestResultsForTestManagmentObject(boundElement, objectQuery, expandSharedSteps);
                                boundElement.AddCustomObjects(replacement.PropertyToEvaluate, value);
                            }
                            else if (objectQuery.DestinationElements.Any(x => x.ItemName.Equals("Configurations")))
                            {
                                var value = GetLinkedConfigurationForTestManagementObject(boundElement, objectQuery, expandSharedSteps);
                                boundElement.AddCustomObjects(replacement.PropertyToEvaluate, value);
                            }
                            else if (objectQuery.DestinationElements.Any(x => x.ItemName.Equals("Test Points")))
                            {
                                var value = GetTestPointsForTestManagmentObject(boundElement, objectQuery, expandSharedSteps);
                                boundElement.AddCustomObjects(replacement.PropertyToEvaluate, value);
                            }
                            else
                            {
                                //Replace the property of the current object with list of workItems
                                var value = GetLinkedWorkItemsForTestManagementObject(boundElement, objectQuery);
                                boundElement.AddCustomObjects(replacement.PropertyToEvaluate, value);
                            }
                        }
                    }

                    if (Cancellation != null && Cancellation())
                    {
                        throw new OperationCanceledException(Resources.TestResult_OperationCancelled);
                    }

                    if (string.IsNullOrEmpty(replacement.LinkedTemplate))
                    {
                        ReplaceSimple(adapter, replacement, boundElement, testAction);
                    }
                    else
                    {
                        ReplaceLinkedTemplate(adapter, replacement, boundElement);
                    }
                }
                catch (Exception ce)
                {
                    adapter.ReleaseNestedAdapter();
                    var bookmarkName = TestReportAdapter.AddErrorBookmark(replacement.Bookmark);
                    var id = string.Empty;
                    try
                    {
                        // Info: double try-cacht necessary for eaxample for WorkItemViewerLink: because no Id-Element exists for the Property "Links", otherwise exception is raised here and jumped into a higher exception block.
                        id = boundElement != null ? boundElement.PropertyValue("Id", (x => x)) : string.Empty;
                    }
                    catch (ConfigurationException configurationException)
                    {
                        id = boundElement.AssociatedObject.GetType().ToString();
                    }
                    var propertyToEvaluate = replacement.PropertyToEvaluate;
                    var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
                    infoStorage.AddItem(new UserInformation
                    {
                        Text = ce.Message + $" - ID: {id}; Property: {propertyToEvaluate}",
                        Explanation = ce.InnerException == null ? string.Empty : ce.InnerException.Message,
                        NavigateToSourceAction = () => TestReportAdapter.Bookmarks[bookmarkName].Select(),
                        Type = UserInformationType.Warning
                    });
                }
            }
        }

        /// <summary>
        /// Get all linked test cases for a TestManagment object
        /// </summary>
        /// <param name="boundElement"></param>
        /// <param name="objectQuery"></param>
        /// <param name="expandSharedSteps">Defines if shared steps shall be expanded.</param>
        /// <returns></returns>
        private IList<ITfsTestCaseDetail> GetLinkedTestCasesForTestManagmentObject(ITfsPropertyValueProvider boundElement, IObjectQuery objectQuery, bool expandSharedSteps)
        {
            IList<ITfsTestCaseDetail> returnList = null;
            SyncServiceTrace.W(Resources.GetLinkedTestCasesForTestManagmentObject);
            // ReSharper disable CanBeReplacedWithTryCastAndCheckForNull

            if (boundElement is ITfsTestPlanDetail)
            {
                returnList = TfsTestAdapter.GetAllTestCases(((ITfsTestPlanDetail)boundElement).TestPlan.RootTestSuite, expandSharedSteps);
            }

            if (boundElement is ITfsTestSuiteDetail)
            {
                returnList = TfsTestAdapter.GetAllTestCases(((ITfsTestSuiteDetail)boundElement).TestSuite, expandSharedSteps);
            }

            if (boundElement is ITfsTestCaseDetail)
            {
                returnList = new List<ITfsTestCaseDetail>();
                returnList.Add((ITfsTestCaseDetail)boundElement);
            }
            // ReSharper restore CanBeReplacedWithTryCastAndCheckForNull
            return returnList;
        }

        /// <summary>
        /// Get all builds for test managment object
        /// </summary>
        /// <param name="boundElement"></param>
        /// <param name="objectQuery"></param>
        /// <param name="expandSharedSteps">Defines if shared steps should be expanded.</param>
        /// <returns></returns>
        private object GetBuildsForTestManagmentObject(ITfsPropertyValueProvider boundElement, IObjectQuery objectQuery, bool expandSharedSteps)
        {
            var returnList = new List<IBuildDetail>();
            //Get all testcases
            SyncServiceTrace.W(Resources.GetBuildsForTestManagmentObject);
            var testCases = GetLinkedTestCasesForTestManagmentObject(boundElement, objectQuery, expandSharedSteps);

            if (testCases == null)
            {
                return returnList;
            }
            //Get the list of builds
            var reloadBuilds = true;
            foreach (var tfsTestCaseDetail in testCases)
            {
                IEnumerable<IBuildDetail> buildForTestCase;
                if (IncludeMostRecentTestResults || IncludeOnlyMostRecentTestResultForAllConfigurations)
                {
                    buildForTestCase = TfsTestAdapter.GetBuildsForTestCase(tfsTestCaseDetail, IncludeMostRecentTestResults, IncludeOnlyMostRecentTestResultForAllConfigurations, CurrentTestConfiguration, reloadBuilds);
                }
                else
                {
                    buildForTestCase = TfsTestAdapter.GetAllBuildsForTestCase(tfsTestCaseDetail, reloadBuilds);
                }

                foreach (var b in buildForTestCase)
                {
                    if (!returnList.Any(x => x.BuildNumber.Equals(b.BuildNumber)))
                    {
                        returnList.Add(b);
                    }
                }
                reloadBuilds = false;
            }
            return returnList;
        }

        /// <summary>
        /// Get all linked configurations for the given test managment object
        /// </summary>
        /// <param name="boundElement"></param>
        /// <param name="objectQuery"></param>
        /// <param name="expandSharedSteps"></param>
        /// <returns></returns>
        public IList<ITfsPropertyValueProvider> GetLinkedConfigurationForTestManagementObject(ITfsPropertyValueProvider boundElement, IObjectQuery objectQuery, bool expandSharedSteps)
        {
            var returnList = new List<ITfsPropertyValueProvider>();
            SyncServiceTrace.W(Resources.GetLinkedConfigurationForTestManagementObject);
            //Configurations are always attached to testcases --> Get all test cases
            IList<ITfsTestCaseDetail> testCases = new List<ITfsTestCaseDetail>();
            ITfsTestPlan testPlan = null;
            ITfsTestSuite testSuite = null;
            // ReSharper disable CanBeReplacedWithTryCastAndCheckForNull

            if (boundElement is ITfsTestPlanDetail)
            {
                testCases = TfsTestAdapter.GetAllTestCases(((ITfsTestPlanDetail)boundElement).TestPlan.RootTestSuite, expandSharedSteps);
                testPlan = ((ITfsTestPlanDetail)boundElement).TestPlan;
            }

            if (boundElement is ITfsTestSuiteDetail)
            {
                testCases = TfsTestAdapter.GetAllTestCases(((ITfsTestSuiteDetail)boundElement).TestSuite, expandSharedSteps);
                testSuite = ((ITfsTestSuiteDetail)boundElement).TestSuite;
                testPlan = ((ITfsTestSuiteDetail)boundElement).TestSuite.AssociatedTestPlan;
            }

            if (boundElement is ITfsTestCaseDetail)
            {
                testCases = new List<ITfsTestCaseDetail>();
                testCases.Add((ITfsTestCaseDetail)boundElement);
                testPlan = ((ITfsTestCaseDetail)boundElement).TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan;
                // todo: verify impact
                testSuite = ((ITfsTestCaseDetail)boundElement).TestCase.AssociatedTestSuiteDetail;
            }

            // ReSharper restore CanBeReplacedWithTryCastAndCheckForNull
            //Get all configurations for the testcases

            if (testPlan == null || testCases == null)
            {
                return returnList;
            }

            var configurations = TfsTestAdapter.GetAssignedTestConfigurationsForTestCases(testPlan, testCases, IncludeMostRecentTestResults, IncludeOnlyMostRecentTestResultForAllConfigurations, testSuite);

            if (objectQuery.FilterOption != null)
            {
                returnList.AddRange(FilterResults(objectQuery, configurations));
            }
            else
            {
                returnList.AddRange(configurations);
            }

            return returnList;
        }

        /// <summary>
        /// Get all linked test points for the given test managment object
        /// </summary>
        /// <param name="boundElement"></param>
        /// <param name="objectQuery"></param>
        /// <param name="expandSharedSteps">Defines if shared steps should be expanded.</param>
        /// <returns></returns>
        private IList<ITfsPropertyValueProvider> GetTestPointsForTestManagmentObject(ITfsPropertyValueProvider boundElement, IObjectQuery objectQuery, bool expandSharedSteps)
        {
            var returnList = new List<ITfsPropertyValueProvider>();
            //Configurations are always attached to testcases --> Get all test cases
            ITfsTestPlan testPlan;

            // ReSharper disable CanBeReplacedWithTryCastAndCheckForNull
            var testPoints = new List<ITfsTestPointDetail>();
            if (boundElement is ITfsTestPlanDetail)
            {
                testPlan = ((ITfsTestPlanDetail)boundElement).TestPlan;

                //Get all configurations for the testcases
                testPoints = (List<ITfsTestPointDetail>)TfsTestAdapter.GetTestPointsForTestPlan(testPlan, boundElement);
            }

            if (boundElement is ITfsTestSuiteDetail)
            {
                var testCases = TfsTestAdapter.GetAllTestCases(((ITfsTestSuiteDetail)boundElement).TestSuite, expandSharedSteps);
                testPlan = ((ITfsTestSuiteDetail)boundElement).TestSuite.AssociatedTestPlan;

                //Get all configurations for the testcases
                testPoints = (List<ITfsTestPointDetail>)TfsTestAdapter.GetTestPointsForTestCases(testPlan, testCases);
            }

            if (boundElement is ITfsTestCaseDetail)
            {
                throw new ConfigurationException("Querying TestPoints for TestCases is currently not supported");
            }
            // ReSharper restore CanBeReplacedWithTryCastAndCheckForNull

            //Filter the list
            returnList.AddRange(FilterResults(objectQuery, testPoints));

            return returnList;
        }

        /// <summary>
        /// Filte the results based on the settings of an object query
        /// </summary>
        /// <param name="objectQuery"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        private static IEnumerable<ITfsPropertyValueProvider> FilterResults(IObjectQuery objectQuery, IEnumerable<ITfsPropertyValueProvider> values)
        {
            Guard.ThrowOnArgumentNull(objectQuery, "ObjectQuery");
            Guard.ThrowOnArgumentNull(values, "values");

            //The list that is returned. Add all values to it
            var returnList = new List<ITfsPropertyValueProvider>();
            returnList.AddRange(values);

            if (objectQuery.FilterOption == null)
            {
                return returnList;
            }

            //The working list for filtering
            var workingList = new List<ITfsPropertyValueProvider>(values);

            string filterProperty;

            if (objectQuery.FilterOption.FilterProperty == null)
            {
                filterProperty = "Id";
            }
            else
            {
                filterProperty = objectQuery.FilterOption.FilterProperty;
            }

            if (objectQuery.FilterOption == null)
            {
                return values;
            }

            //Filter by latest --> Group by id, then by date completed
            if (objectQuery.FilterOption.Latest)
            {
                //Group by the Testcase id
                workingList = (List<ITfsPropertyValueProvider>)workingList.GroupBy(i => i.PropertyValue(filterProperty, x => x)).SelectMany(g => g.Where(i => i.PropertyValue("DateCompleted", x => x) == g.Max(m => m.PropertyValue("DateCompleted", x => x)))).Distinct();
                returnList = workingList;
            }

            if (objectQuery.FilterOption.Distinct)
            {
                //var property = objectQuery.FilterOption.FilterProperty;
                workingList = workingList.GroupBy(p => p.PropertyValue(filterProperty, x => x)).Select(g => g.First()).ToList();
                returnList = workingList;
            }

            return returnList;
        }

        /// <summary>
        /// Get all linked test results for a testmanagment object
        /// </summary>
        /// <param name="boundElement"></param>
        /// <param name="objectQuery"></param>
        /// <param name="expandSharedSteps">Defines if shared steps should be expanded.</param>
        /// <returns></returns>
        private IList<ITfsPropertyValueProvider> GetLinkedTestResultsForTestManagmentObject(ITfsPropertyValueProvider boundElement, IObjectQuery objectQuery, bool expandSharedSteps)
        {
            SyncServiceTrace.W(Resources.GetLinkedTestResultsForTestManagmentObject);
            var returnList = new List<ITfsPropertyValueProvider>();
            IList<ITfsTestCaseDetail> testCases = new List<ITfsTestCaseDetail>();

            // ReSharper disable CanBeReplacedWithTryCastAndCheckForNull
            if (boundElement is ITfsTestPlanDetail)
            {
                //Get all TestCases
                testCases = TfsTestAdapter.GetAllTestCases(((ITfsTestPlanDetail)boundElement).TestPlan.RootTestSuite, expandSharedSteps);
            }

            if (boundElement is ITfsTestSuiteDetail)
            {
                //Get all TestCases
                testCases = TfsTestAdapter.GetAllTestCases(((ITfsTestSuiteDetail)boundElement).TestSuite, expandSharedSteps);
            }

            if (boundElement is ITfsTestCaseDetail)
            {
                testCases = new List<ITfsTestCaseDetail>();
                testCases.Add((ITfsTestCaseDetail)boundElement);
            }
            // ReSharper restore CanBeReplacedWithTryCastAndCheckForNull

            //Get All TestResults for all TestCases
            foreach (var tc in testCases)
            {
                IEnumerable<ITfsTestResultDetail> testResults;
                if (IncludeMostRecentTestResults)
                {
                    //Get the last testresult for each configuration
                    if (IncludeOnlyMostRecentTestResultForAllConfigurations)
                    {
                        testResults = TfsTestAdapter.GetLatestTestResultsForAllSelectedConfigurations(tc.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tc.TestCase, CurrentTestConfiguration, null, tc.TestCase.AssociatedTestSuiteDetail, false);
                    }
                    else
                    {
                        testResults = TfsTestAdapter.GetLatestTestResults(tc.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tc.TestCase, null, null, tc.TestCase.AssociatedTestSuiteDetail, false);
                    }
                }
                else
                {
                    Dictionary<int, int> lastTestRunPerConfig;
                    testResults = TfsTestAdapter.GetTestResults(tc.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tc.TestCase, CurrentTestConfiguration, null, tc.TestCase.AssociatedTestSuiteDetail, false, out lastTestRunPerConfig);
                }
                returnList.AddRange(testResults);
            }

            returnList = FilterTestResults(objectQuery, returnList);
            return returnList;
        }

        /// <summary>
        /// Filter a given list of Test Results based on criteria provided by an object query
        /// </summary>
        /// <param name="objectQuery"></param>
        /// <param name="returnList"></param>
        /// <returns></returns>
        private List<ITfsPropertyValueProvider> FilterTestResults(IObjectQuery objectQuery, List<ITfsPropertyValueProvider> returnList)
        {
            if (returnList == null || returnList.Count == 0)
            {
                return returnList;
            }

            //Filter the results based on the object qurey
            if (objectQuery.FilterOption != null)
            {
                SyncServiceTrace.W($"Filter Test Results: { objectQuery.ToString()}");
                var originalSize = returnList.Count;

                var newReturnList = new List<ITfsPropertyValueProvider>(returnList);

                if (objectQuery.FilterOption.Latest)
                {
                    var propertyToFilterLatest = "TestCaseId";
                    if (objectQuery.FilterOption.FilterProperty != null)
                    {
                        propertyToFilterLatest = objectQuery.FilterOption.FilterProperty;
                    }

                    //Group by the Testcase id
                    var newList = newReturnList.GroupBy(i => i.PropertyValue(propertyToFilterLatest, x => x)).SelectMany(g => g.Where(i => i.PropertyValue("DateCompleted", x => x) == g.Max(m => m.PropertyValue("DateCompleted", x => x)))).Distinct();

                    newReturnList = new List<ITfsPropertyValueProvider>(newList);
                    returnList = newReturnList;
                }

                if (objectQuery.FilterOption.Distinct)
                {
                    var property = "TestCaseId";
                    if (objectQuery.FilterOption.FilterProperty != null)
                    {
                        property = objectQuery.FilterOption.FilterProperty;
                    }
                    var filteredList = newReturnList.GroupBy(p => p.PropertyValue(property, x => x)).Select(g => g.First()).ToList();
                    returnList = filteredList;
                }
                SyncServiceTrace.W(string.Format(Resources.FilterTestResults, originalSize, returnList.Count));
            }

            return returnList;
        }

        /// <summary>
        /// Filter a given list of work items based on criteria provided by an object query
        /// </summary>
        /// <param name="objectQuery"></param>
        /// <param name="returnList"></param>
        /// <returns></returns>
        private List<WorkItem> FilterWorkItems(IObjectQuery objectQuery, List<WorkItem> returnList)
        {
            if (returnList == null || returnList.Count == 0)
            {
                return returnList;
            }

            //Filter the results based on the object qurey
            if (objectQuery.FilterOption != null)
            {
                SyncServiceTrace.W($"Filter Workitems: {objectQuery.ToString()}");

                if (objectQuery.FilterOption.Latest)
                {
                    throw new ConfigurationException("The filter option latest cannot be used for linked work items");
                }
                if (objectQuery.FilterOption.FilterProperty != null)
                {
                    throw new ConfigurationException("Filter properties cannot be used for linked work items");
                }

                var newReturnList = new List<WorkItem>(returnList);

                if (objectQuery.FilterOption.Distinct)
                {
                    var filteredList = newReturnList.GroupBy(p => p.Id).Select(g => g.First()).ToList();

                    SyncServiceTrace.W(string.Format(Resources.FilterWorkItems, returnList.Count, filteredList.Count));

                    returnList = filteredList;
                }
            }
            return returnList;
        }

        /// <summary>
        /// Get specific workItems based on a filter criteria
        /// </summary>
        /// <param name="boundElement"></param>
        /// <param name="objectQuery"></param>
        /// <returns></returns>
        private List<WorkItem> GetLinkedWorkItemsForTestManagementObject(ITfsPropertyValueProvider boundElement, IObjectQuery objectQuery)
        {
            Guard.ThrowOnArgumentNull(objectQuery, "ObjectQuery");
            Guard.ThrowOnArgumentNull(boundElement, "Element");

            SyncServiceTrace.W(Resources.GetLinkedWorkItemsForTestManagementObject);

            //Build the filter criteria
            var workItemTypes = new List<string>();
            foreach (var workItemType in objectQuery.DestinationElements)
            {
                workItemTypes.Add(workItemType.ItemName);
            }

            //Determine type of object
            //TestPlan --> Get All TestCases
            //TestSuite --> Get All Testcases
            //Testcase --> Search the linked items directly
            IList<ITfsTestCaseDetail> testCases = null;
            var returnList = new List<WorkItem>();
            // ReSharper disable CanBeReplacedWithTryCastAndCheckForNull

            var expandSharedSteps = Configuration.ConfigurationTest.ExpandSharedSteps;

            //Test Results need a different approach
            if (boundElement is ITfsTestResultDetail)
            {
                var allWorkItemsForTestResult = TfsTestAdapter.GetAllLinkedWorkItemsForTestResult((ITfsTestResultDetail)boundElement, workItemTypes, objectQuery.WorkItemLinkFilters, objectQuery.FilterOption);
                if (allWorkItemsForTestResult != null)
                {
                    returnList.AddRange(allWorkItemsForTestResult);
                }
            }

            if (boundElement is ITfsTestPlanDetail)
            {
                //Get all TestCases
                testCases = TfsTestAdapter.GetAllTestCases(((ITfsTestPlanDetail)boundElement).TestPlan.RootTestSuite, expandSharedSteps);
            }

            if (boundElement is ITfsTestSuiteDetail)
            {
                //Get all TestCases
                testCases = TfsTestAdapter.GetAllTestCases(((ITfsTestSuiteDetail)boundElement).TestSuite, expandSharedSteps);
            }

            if (boundElement is ITfsTestCaseDetail)
            {
                testCases = new List<ITfsTestCaseDetail>();
                testCases.Add((ITfsTestCaseDetail)boundElement);
            }

            if (testCases != null)
            {
                var allWorkItemsForTestCases = TfsTestAdapter.GetAllLinkedWorkItemsForTestCases(testCases, workItemTypes, objectQuery.WorkItemLinkFilters, objectQuery.FilterOption);
                returnList.AddRange(allWorkItemsForTestCases);
                //Get All TestResults for all TestCases
                foreach (var tc in testCases)
                {
                    IEnumerable<ITfsTestResultDetail> testResults;
                    if (IncludeMostRecentTestResults)
                    {
                        //Get the last testresult for each configuration
                        if (IncludeOnlyMostRecentTestResultForAllConfigurations)
                        {
                            testResults = TfsTestAdapter.GetLatestTestResultsForAllSelectedConfigurations(tc.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tc.TestCase, CurrentTestConfiguration, null, tc.TestCase.AssociatedTestSuiteDetail, true);
                        }
                        else
                        {
                            testResults = TfsTestAdapter.GetLatestTestResults(tc.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tc.TestCase, null, null, tc.TestCase.AssociatedTestSuiteDetail, true);
                        }
                    }
                    else
                    {
                        Dictionary<int, int> lastTestRunPerConfig;
                        testResults = TfsTestAdapter.GetTestResults(tc.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tc.TestCase, CurrentTestConfiguration, null, tc.TestCase.AssociatedTestSuiteDetail, true, out lastTestRunPerConfig);
                    }

                    if (testResults != null)
                    {
                        foreach (var tfsTestResultDetail in testResults)
                        {
                            var allWorkItemsForTestResult = TfsTestAdapter.GetAllLinkedWorkItemsForTestResult(tfsTestResultDetail, workItemTypes, objectQuery.WorkItemLinkFilters, objectQuery.FilterOption);
                            returnList.AddRange(allWorkItemsForTestResult);
                        }
                    }
                }
            }

            //Filter the list

            returnList = FilterWorkItems(objectQuery, returnList);
            return returnList;
        }

        /// <summary>
        /// Get all parameters for all iterations for a testcase
        /// </summary>
        /// <param name="boundElement"></param>
        /// <returns></returns>
        private static ITestCaseParameters GetParametersForTestCase(ITfsPropertyValueProvider boundElement)
        {
            //Cast the bound element to Testcase
            var testCase = (ITestCase)boundElement.AssociatedObject;
            var result = new TestCaseParameters();

            //Build the header 
            var allParameters = "";
            foreach (DataColumn col in testCase.DefaultTable.Columns)
            {
                allParameters += col.Caption;
                allParameters += ";";
            }
            var i = 1;
            foreach (DataRow row in testCase.DefaultTable.Rows)
            {
                //Collect the data for each iteration
                var allParametersforIterations = new List<Parameter>();

                var tci = new TestCaseIterations();
                tci.IterationNumber = i;
                foreach (DataColumn col in testCase.DefaultTable.Columns)
                {
                    var currentParameter = "";
                    var para = new Parameter();

                    currentParameter += col.Caption;
                    currentParameter += "=";
                    currentParameter += row[col];

                    para.ParameterValue = currentParameter;
                    allParametersforIterations.Add(para);
                }
                tci.Parameters = allParametersforIterations;
                result.Iterations.Add(tci);
                result.AllParameters = allParameters;
                ++i;
            }
            return result;
        }

        private void ReplaceLinkedTemplate(IWord2007TestReportAdapter adapter, IConfigurationTestReplacement replacement, ITfsPropertyValueProvider valueProvider)
        {
            if (adapter == null)
            {
                throw new ArgumentNullException("adapter");
            }
            if (replacement == null)
            {
                throw new ArgumentNullException("replacement");
            }
            if (valueProvider == null)
            {
                throw new ArgumentNullException("valueProvider");
            }

            var firstTimeReplaced = true;
            // Linked template replacement
            // Call valueProvider.PropertyValue(replacement.Property, (value) => 'add value')
            valueProvider.PropertyValue(replacement.PropertyToEvaluate, value =>
                {
                    if (Cancellation != null && Cancellation())
                    {
                        throw new OperationCanceledException(Resources.TestResult_OperationCancelled);
                    }

                    // Create copy of template - temporary template file
                    var templateName = replacement.LinkedTemplate;
                    // Get the conditions
                    var conditions = Configuration.ConfigurationTest.GetConditions(templateName);
                    // Iterate the conditions and
                    foreach (var condition in conditions)
                    {
                        if (string.IsNullOrEmpty(condition.PropertyToEvaluate) || condition.Values == null || condition.Values.Count == 0)
                        {
                            continue;
                        }
                        var nestedProvider = valueProvider.GetTemporaryPropertyValueProvider(value);
                        var valueOfNestedProvider = nestedProvider.PropertyValue(condition.PropertyToEvaluate, EnumerableExpander);
                        if (condition.Values.Contains(valueOfNestedProvider) == false)
                        {
                            // Condition is evaluated to false, dont insert anything
                            return;
                        }
                    }

                    // Get temporary file
                    var temporaryTemplateFile = TempFolder.GetTemporaryFile();
                    // Get source file
                    var sourceFile = Configuration.ConfigurationTest.GetFileName(templateName);
                    // Copy the template to temp file
                    File.Copy(sourceFile, temporaryTemplateFile);

                    IWord2007TestReportAdapter nestedAdapter = null;
                    try
                    {
                        // Open temporary template file
                        nestedAdapter = adapter.CreateNestedAdapter(temporaryTemplateFile);

                        // Get replacements - used in 'callback' action
                        var replacements = Configuration.ConfigurationTest.GetReplacements(templateName);

                        // Obtained value is object with next properties defined in linked template
                        // If we are currently replacing a test action, check if this is the first action step.
                        var testAction = value as TfsTestAction;
                        var expandSharedSteps = Configuration.ConfigurationTest.ExpandSharedSteps;
                        Replace(nestedAdapter, replacements, valueProvider.GetTemporaryPropertyValueProvider(value), expandSharedSteps, testAction);
                    }
                    finally
                    {
                        // Close temporary template file
                        if (nestedAdapter != null)
                        {
                            nestedAdapter.ReleaseNestedAdapter();
                        }
                    }

                    if (firstTimeReplaced)
                    {
                        firstTimeReplaced = false;
                        var headerTemplateName = Configuration.ConfigurationTest.GetHeaderTemplate(templateName);
                        if (!string.IsNullOrEmpty(headerTemplateName))
                        {
                            adapter.InsertFile(replacement.Bookmark, Configuration.ConfigurationTest.GetFileName(headerTemplateName), Configuration.ConfigurationTest.GetPreOperations(headerTemplateName), Configuration.ConfigurationTest.GetPostOperations(headerTemplateName));
                        }
                    }

                    // Replace bookmark
                    adapter.InsertFile(replacement.Bookmark, temporaryTemplateFile, Configuration.ConfigurationTest.GetPreOperations(templateName), Configuration.ConfigurationTest.GetPostOperations(templateName));

                    // Delete used temporary template file
                    File.Delete(temporaryTemplateFile);
                }, EnumerableExpander);

            // Clean up after linked template replacement
            adapter.ReplaceBookmarkText(replacement.Bookmark, string.Empty, PropertyValueFormat.PlainText, replacement.WordBookmark);
        }

        private void ReplaceSimple(IWord2007TestReportAdapter adapter, IConfigurationTestReplacement replacement, ITfsPropertyValueProvider valueProvider, TfsTestAction testAction = null)
        {
            Guard.ThrowOnArgumentNull(adapter, "adapter");
            Guard.ThrowOnArgumentNull(replacement, "replacement");
            Guard.ThrowOnArgumentNull(valueProvider, "valueProvider");

            if (replacement.UriLink != null)
            {
                // insert hyperlink to web view of the work item
                var value = valueProvider.PropertyValue(replacement.PropertyToEvaluate, EnumerableExpander);
                var uri = valueProvider.PropertyValue(replacement.UriLink.Uri, EnumerableExpander);
                adapter.ReplaceBookmarkHyperlink(replacement.Bookmark, value, uri);
            }
            else if (replacement.AttachmentLink != null)
            {
                try
                {
                    ReplaceAttachment(adapter, replacement, valueProvider);
                }
                catch (AttachmentException ae)
                {
                    var bookmarkName = TestReportAdapter.AddErrorBookmark(replacement.Bookmark);
                    var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
                    var id = valueProvider != null ? valueProvider.PropertyValue("Id", (x => x)) : string.Empty;
                    var propertyToEvaluate = replacement.PropertyToEvaluate;
                    infoStorage.AddItem(new UserInformation { Text = ae.Message + string.Format(CultureInfo.InvariantCulture, " - Attachment ID: {0}; Property to evaluate: {1}", id, propertyToEvaluate), Explanation = ae.InnerException == null ? string.Empty : ae.InnerException.Message, NavigateToSourceAction = () => TestReportAdapter.Bookmarks[bookmarkName].Select(), Type = UserInformationType.Warning });
                }
            }
            else if (replacement.WorkItemEditorLink != null)
            {
                var value = valueProvider.PropertyValue(replacement.PropertyToEvaluate, EnumerableExpander);
                var artifactUri = valueProvider.PropertyValue(replacement.WorkItemEditorLink.Uri, EnumerableExpander);
                var workItemEditorLink = TfsTestAdapter.GetWorkItemEditorLink(artifactUri).AbsoluteUri;
                adapter.ReplaceBookmarkHyperlink(replacement.Bookmark, value, workItemEditorLink);
            }
            else if (replacement.WorkItemViewerLink != null)
            {
                var value = valueProvider.PropertyValue(replacement.PropertyToEvaluate, EnumerableExpander);
                var valueOfIdField = valueProvider.PropertyValue(replacement.WorkItemViewerLink.Id, EnumerableExpander);
                var id = int.Parse(valueOfIdField, CultureInfo.CurrentCulture);

                var revision = -1;
                if (!string.IsNullOrEmpty(replacement.WorkItemViewerLink.Revision))
                {
                    var revisionValue = valueProvider.PropertyValue(replacement.WorkItemViewerLink.Revision, EnumerableExpander);
                    revision = int.Parse(revisionValue, CultureInfo.CurrentCulture);
                }

                if (replacement.WorkItemViewerLink.AutoText)
                {
                    var title = TfsTestAdapter.GetWorkItemTitle(id);
                    if (title != null)
                    {
                        value = string.Format(CultureInfo.CurrentCulture, Resources.TestReportHelper_WorkItemViewLink_AutoText, id, title);
                    }
                }

                if (!string.IsNullOrEmpty(replacement.WorkItemViewerLink.Format))
                {
                    var myTfsWorkItem = TfsTestAdapter.GetTfsWorkItemById(id);
                    var formattedWorkitem = WorkItemStringFormatter.GetWorkItemFomatted(myTfsWorkItem, replacement.WorkItemViewerLink.Format);
                    value = formattedWorkitem;
                }

                var workItemViewerLink = TfsTestAdapter.GetWorkItemViewerLink(id, revision).AbsoluteUri;
                adapter.ReplaceBookmarkHyperlink(replacement.Bookmark, value, workItemViewerLink);
            }
            else if (replacement.BuildViewerLink != null)
            {
                var value = valueProvider.PropertyValue(replacement.PropertyToEvaluate, EnumerableExpander);
                var buildNumber = valueProvider.PropertyValue(replacement.BuildViewerLink.BuildNumber, EnumerableExpander);

                // Test result may be started without build. In this case is no build number in test result.
                if (string.IsNullOrEmpty(buildNumber))
                {
                    adapter.ReplaceBookmarkText(replacement.Bookmark, value, replacement.ValueType, replacement.WordBookmark);
                }
                else
                {
                    var buildViewerLink = TfsTestAdapter.GetBuildViewerLink(buildNumber)?.AbsoluteUri;
                    if (buildViewerLink != null)
                    {
                        adapter.ReplaceBookmarkHyperlink(replacement.Bookmark, value, buildViewerLink);
                    }
                }
            }
            else if (valueProvider.GetCustomObject(replacement.PropertyToEvaluate) != null)
            {
                //Replace a custom object
                var value = valueProvider.GetCustomObject(replacement.PropertyToEvaluate);
                adapter.ReplaceBookmarkText(replacement.Bookmark, value.ToString(), replacement.ValueType, replacement.WordBookmark);
            }
            else
            {
                // Simple replacement
                var value = valueProvider.PropertyValue(replacement.PropertyToEvaluate, EnumerableExpander);
                int intValue = 0;
                if (replacement.PropertyToEvaluate == "ResolutionStateId" && replacement.ResolveResolutionState && Int32.TryParse(value, out intValue))
                {
                    value = _resolutionStates[intValue];
                }

                if (replacement.PropertyToEvaluate == "FailureType" && value == "NullValue")
                {
                    value = string.Empty;
                }

                if (testAction != null)
                {
                    value = HtmlHelper.ReplaceParagraphAndDivisionTags(value);

                    // Use HTML to mark shared steps bold
                    if (testAction.IsSharedStepTitle)
                    {
                        adapter.ReplaceBookmarkText(replacement.Bookmark, value, PropertyValueFormat.HTMLBold, replacement.WordBookmark);
                    }
                    else
                    {
                        adapter.ReplaceBookmarkText(replacement.Bookmark, value, replacement.ValueType, replacement.WordBookmark);
                    }
                }
                else
                {
                    adapter.ReplaceBookmarkText(replacement.Bookmark, value, replacement.ValueType, replacement.WordBookmark);
                }
            }
        }

        private void ReplaceAttachment(IWord2007TestReportAdapter adapter, IConfigurationTestReplacement replacement, ITfsPropertyValueProvider valueProvider)
        {
            var attachment = valueProvider.AssociatedObject as ITestAttachment;
            if (attachment == null)
            {
                throw new ConfigurationException("The AttachmentLink-Element can only be used on properties of an object of type ITestAttachment.");
            }
            //Dont download the attachment if it is not complete or the lenth is 0
            if (!attachment.IsComplete || attachment.Length == 0)
            {
                var attachmentName = attachment.Name;
                throw new AttachmentException($"The uploaded attachment {attachmentName} is inclomplete or corrupt. It has not been downloaded.");
            }

            // We want to save attachments in the same folder as the reporting document. For this
            // we need to get a hold of the adapter that works with the reporting document and not
            // use the provided adapter that is most likely nested. 
            Func<IWord2007TestReportAdapter, IWord2007TestReportAdapter> getRoot = null;
            getRoot = x => x.ParentAdapter == null ? x : getRoot(x.ParentAdapter);
            var root = getRoot(adapter);

            var isFirstAttachment = SaveDocumentAndEnsureHyperlinkBase(root);

            if (replacement.AttachmentLink.Mode != AttachmentLinkMode.LinkToServerVersion)
            {
                var filename = string.Empty;
                var relativePathToAttachment = string.Empty;

                if (Configuration.AttachmentFolderMode == AttachmentFolderMode.BasedOnTestSuite)
                {
                    HandleTestSuiteHierarchyForAttachments(root, attachment, isFirstAttachment, out filename, out relativePathToAttachment);
                }
                else
                {
                    HandleGuidOrWithoutGuidModeForAttachments(root, attachment, valueProvider, out filename, out relativePathToAttachment);
                }

                // simulate properties "LocalPath" and "LocalFilename"
                string propertyValue;
                switch (replacement.PropertyToEvaluate)
                {
                    case "LocalPath":
                        propertyValue = relativePathToAttachment;
                        break;

                    case "LocalFilename":
                        propertyValue = filename;
                        break;

                    default:
                        propertyValue = valueProvider.PropertyValue(replacement.PropertyToEvaluate, EnumerableExpander);
                        break;
                }

                // insert as normal text or hyperlink to local file
                if (replacement.AttachmentLink.Mode == AttachmentLinkMode.DownloadAndLinkToLocalFile || replacement.AttachmentLink.Mode == AttachmentLinkMode.DownloadAndLinkToLocalFileWithoutGuid)
                {
                    adapter.ReplaceBookmarkHyperlink(replacement.Bookmark, propertyValue, relativePathToAttachment);
                }
                else
                {
                    adapter.ReplaceBookmarkText(replacement.Bookmark, propertyValue, PropertyValueFormat.PlainText, replacement.WordBookmark);
                }
            }
            else
            {
                // Link to attachment on tfs server
                adapter.ReplaceBookmarkHyperlink(replacement.Bookmark, valueProvider.PropertyValue(replacement.PropertyToEvaluate, EnumerableExpander), attachment.Uri.ToString());
            }
        }

        private void HandleGuidOrWithoutGuidModeForAttachments(IWord2007TestReportAdapter root, ITestAttachment attachment, ITfsPropertyValueProvider valueProvider, out string filename, out string relativePathToAttachment)
        {
            var guid = Guid.NewGuid().ToString();
            var name = valueProvider.PropertyValue("Name", EnumerableExpander);
            filename = string.Format(CultureInfo.InvariantCulture, Resources.TestReportHelper_AttachmentFilename, guid, name);
            relativePathToAttachment = Path.Combine(root.AttachmentFolder, filename);
            var aboluteAttachmentFolderPath = Path.Combine(root.DocumentPath, root.AttachmentFolder);
            var absolutePath = Path.Combine(aboluteAttachmentFolderPath, filename);
            var absolutePathDirName = Path.GetDirectoryName(absolutePath) ?? root.DocumentPath;
            if (!Directory.Exists(absolutePathDirName))
            {
                Directory.CreateDirectory(absolutePathDirName);
            }

            attachment.DownloadToFile(absolutePath);
        }

        private string FindAboluteAttachmentFolderPath(IWord2007TestReportAdapter root, string testSuiteFolderPath, bool isFirstAttachment)
        {
            var rootFolder = Path.Combine(root.DocumentPath, $"{root.AttachmentFolder}");
            var aboluteAttachmentFolderPath = string.Empty;

            if (Directory.Exists(rootFolder) && isFirstAttachment)
            {
                _isDuplicateFolderCreated = true;
                _guid = Guid.NewGuid().ToString();
                aboluteAttachmentFolderPath = Path.Combine(root.DocumentPath, $"{root.AttachmentFolder}_{_guid}\\{testSuiteFolderPath}");
            }
            else if (_isDuplicateFolderCreated)
            {
                aboluteAttachmentFolderPath = Path.Combine(root.DocumentPath, $"{root.AttachmentFolder}_{_guid}\\{testSuiteFolderPath}");
            }
            else
            {
                aboluteAttachmentFolderPath = Path.Combine($"{rootFolder}\\{testSuiteFolderPath}");
            }

            return aboluteAttachmentFolderPath;

        }


        private void HandleTestSuiteHierarchyForAttachments(IWord2007TestReportAdapter root, ITestAttachment attachment, bool isFirstAttachment, out string filename, out string relativePathToAttachment)
        {
            var testSuiteFolderPath = string.Empty;
            var aboluteAttachmentFolderPath = string.Empty;
            var absolutePath = string.Empty;


            testSuiteFolderPath = $"{_testSuite.TestSuite.Title}_({_testSuite.TestSuite.Id})";
            if (TestResultReportModel.Paths.ContainsKey(_testSuite.TestSuite.Id))
            {
                aboluteAttachmentFolderPath = TestResultReportModel.Paths.Where(x => x.Key == _testSuite.TestSuite.Id).SingleOrDefault().Value;
            }
            else
            {
                aboluteAttachmentFolderPath = FindAboluteAttachmentFolderPath(root, testSuiteFolderPath, isFirstAttachment);
                TestResultReportModel.Paths.Add(_testSuite.TestSuite.Id, aboluteAttachmentFolderPath);
            }

            if (_testSuite.TestSuite.TestSuites.Any())
            {
                foreach (var childTestSuite in _testSuite.TestSuite.TestSuites)
                {
                    var path = $"{aboluteAttachmentFolderPath}\\{childTestSuite.Title}_({childTestSuite.Id})";
                    if (!TestResultReportModel.Paths.ContainsKey(childTestSuite.Id))
                    {
                        TestResultReportModel.Paths.Add(childTestSuite.Id, path);
                    }
                }
            }

            var nameOfAttachment = attachment.Name;
            var exstension = Path.GetExtension(nameOfAttachment);
            var index = nameOfAttachment.Length - exstension.Length;
            var name = nameOfAttachment.Substring(0, index);
            filename = string.Empty;

            var iteration = _testCase.IterationPath;
            var isAttachmentNameGenerated = false;

            if (_testCase.OriginalTestCase.Actions.Any())
            {
                var steps = _testCase.OriginalTestCase.Actions.Where(x => x as ITestStep != null);
                var countOfAllAttachmentsFromAllStepsPerTestCase = 0;
                foreach (var item in steps)
                {
                    var step = item as ITestStep;
                    if (step != null)
                    {
                        countOfAllAttachmentsFromAllStepsPerTestCase += step.Attachments.Count();
                    }
                }

                if (_numberOfStepAttachmentsPerTestCase < countOfAllAttachmentsFromAllStepsPerTestCase)
                {
                    foreach (var action in _testCase.OriginalTestCase.Actions)
                    {
                        var step = action as ITestStep;

                        if (step != null && step.Attachments.Any())
                        {
                            foreach (var att in step.Attachments)
                            {
                                if (att.Name.Equals(nameOfAttachment) && (!_stepsAndAttachmentsPerTestCase.ContainsKey(step.Id) || step.Attachments.Count() > _stepsAndAttachmentsPerTestCase[step.Id].Count()))
                                {
                                    var iterationName = iteration.Substring(iteration.LastIndexOf("\\") + 1);
                                    iterationName = iterationName.Replace(' ', '_');


                                    if (_attachmentNames.Contains($"{name}({TEST_CASE_PREFIX}{_testCase.Id.ToString()}_{iterationName}{STEP_PREFIX}{step.Id.ToString()}){exstension}"))
                                    {
                                        var file = $"{name}({TEST_CASE_PREFIX}{_testCase.Id.ToString()}_{iterationName}{STEP_PREFIX}{step.Id.ToString()}){exstension}";
                                        var counter = _attachmentNames.Where(x => x.Equals(file)).Count();
                                        var newFile = $"{name}({TEST_CASE_PREFIX}{_testCase.Id.ToString()}_{iterationName}{STEP_PREFIX}{step.Id.ToString()})_{(counter).ToString()}{exstension}";
                                        filename = $"\\{newFile}";
                                    }
                                    else
                                    {
                                        var file = $"{name}({TEST_CASE_PREFIX}{_testCase.Id.ToString()}_{iterationName}{STEP_PREFIX}{step.Id.ToString()}){exstension}";
                                        filename = $"\\{file}";
                                    }
                                    isAttachmentNameGenerated = true;
                                    _numberOfStepAttachmentsPerTestCase++;

                                    if (_stepsAndAttachmentsPerTestCase.ContainsKey(step.Id))
                                    {
                                        _stepsAndAttachmentsPerTestCase[step.Id].Add(filename);
                                    }
                                    else
                                    {
                                        List<string> list = new List<string>();
                                        list.Add(filename);
                                        _stepsAndAttachmentsPerTestCase.Add(step.Id, list);
                                    }

                                    _attachmentNames.Add($"{name}({TEST_CASE_PREFIX}{_testCase.Id.ToString()}_{iterationName}{STEP_PREFIX}{step.Id.ToString()}){exstension}");
                                }
                                if (isAttachmentNameGenerated) break;
                            }
                        }
                        if (isAttachmentNameGenerated) break;
                    }
                }
                else if (_testCase.OriginalTestCase.Attachments.Any() && _testCase.OriginalTestCase.Actions.Any() && !isAttachmentNameGenerated)
                {
                    if (_attachmentNames.Contains($"{root.AttachmentFileName(nameOfAttachment, _testCase)}_{nameOfAttachment}"))
                    {
                        var counter = _attachmentNames.Where(x => x.Equals($"{root.AttachmentFileName(nameOfAttachment, _testCase)}_{nameOfAttachment}")).Count();
                        filename = string.Format(CultureInfo.InvariantCulture, Resources.TestReportHelper_AttachmentFilename, root.AttachmentFileName(nameOfAttachment, _testCase), $"{nameOfAttachment.Insert(index, $"_{(counter).ToString()}")}");
                    }
                    else
                    {
                        filename = string.Format(CultureInfo.InvariantCulture, Resources.TestReportHelper_AttachmentFilename, root.AttachmentFileName(nameOfAttachment, _testCase), nameOfAttachment);
                    }
                    _attachmentNames.Add($"{root.AttachmentFileName(nameOfAttachment, _testCase)}_{nameOfAttachment}");
                    isAttachmentNameGenerated = true;
                }
            }

            if (_testCase.OriginalTestCase.Attachments.Any() && !_testCase.OriginalTestCase.Actions.Any() && !isAttachmentNameGenerated)
            {
                foreach (var att in _testCase.OriginalTestCase.Attachments)
                {
                    if (att.Name.Equals(nameOfAttachment))
                    {
                        if (_attachmentNames.Contains($"{root.AttachmentFileName(nameOfAttachment, _testCase)}_{nameOfAttachment}"))
                        {
                            var counter = _attachmentNames.Where(x => x.Equals($"{root.AttachmentFileName(nameOfAttachment, _testCase)}_{nameOfAttachment}")).Count();
                            filename = string.Format(CultureInfo.InvariantCulture, Resources.TestReportHelper_AttachmentFilename, root.AttachmentFileName(nameOfAttachment, _testCase), $"{nameOfAttachment.Insert(index, $"_{(counter).ToString()}")}");
                        }
                        else
                        {
                            filename = string.Format(CultureInfo.InvariantCulture, Resources.TestReportHelper_AttachmentFilename, root.AttachmentFileName(nameOfAttachment, _testCase), nameOfAttachment);
                        }
                        break;
                    }
                }

                _attachmentNames.Add($"{root.AttachmentFileName(nameOfAttachment, _testCase)}_{nameOfAttachment}");
            }

            relativePathToAttachment = Path.Combine($"{root.DocumentPath}\\{root.AttachmentFolder}", $"{testSuiteFolderPath}{filename}");

            if (TestResultReportModel.Paths.ContainsKey(_testSuite.TestSuite.Id))
            {
                relativePathToAttachment = $"{TestResultReportModel.Paths.Where(x => x.Key == _testSuite.TestSuite.Id).SingleOrDefault().Value}{filename}";
            }

            absolutePath = $"{aboluteAttachmentFolderPath}{filename}";

            var absolutePathDirName = Path.GetDirectoryName(absolutePath) ?? root.DocumentPath;
            if (!Directory.Exists(absolutePathDirName))
            {
                Directory.CreateDirectory(absolutePathDirName);
            }

            attachment.DownloadToFile(absolutePath);
        }

        /// <summary>
        /// Saves the document and ensures HyperlinkBase-Property of the document is set. Setting of HyperlinkBase can be turned off in the configuration.
        /// </summary>
        /// <param name="wordAdapter">The word adapter to save.</param>
        /// <exception cref="OperationCanceledException">Thrown if the user aborts saving the document.</exception>
        private bool SaveDocumentAndEnsureHyperlinkBase(IWord2007TestReportAdapter wordAdapter)
        {
            var isFirstAttachment = false;
            if (string.IsNullOrEmpty(wordAdapter.DocumentPath))
            {
                if (Configuration.ConfigurationTest.ShowHyperlinkBaseMessageBoxes == false || DialogResult.Yes == MessageBox.Show(Resources.Attachment_SaveDocumentNowText, Resources.MessageBox_Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, 0))
                {
                    try
                    {
                        isFirstAttachment = true;
                        wordAdapter.SaveDocument();
                    }
                    catch (COMException)
                    {
                        // user aborted save
                        throw new OperationCanceledException(Resources.TestResult_OperationCancelled);
                    }

                    if (Configuration.ConfigurationTest.SetHyperlinkBase)
                    {
                        wordAdapter.HyperlinkBase = wordAdapter.DocumentPath;
                    }
                }
            }
            else
            {
                isFirstAttachment = false;
                if (Configuration.ConfigurationTest.SetHyperlinkBase && ( wordAdapter.HyperlinkBase == null || !wordAdapter.HyperlinkBase.Equals(wordAdapter.DocumentPath)))
                {
                    //If the attachment folder does not exist, or no attachment exist under this path we do not have to ask the user if he wants to change the hyperlinkbase
                    if (!Directory.Exists(Path.Combine(wordAdapter.DocumentPath, wordAdapter.AttachmentFolder)) || !Directory.EnumerateFiles(Path.Combine(wordAdapter.DocumentPath, wordAdapter.AttachmentFolder)).Any())
                    {
                        wordAdapter.HyperlinkBase = wordAdapter.DocumentPath;
                    }
                    else
                    {
                        if (Configuration.ConfigurationTest.ShowHyperlinkBaseMessageBoxes == false || DialogResult.Yes == MessageBox.Show(Resources.Attachment_AdjustHyperlinkBase, Resources.MessageBox_Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, 0))
                        {
                            wordAdapter.HyperlinkBase = wordAdapter.DocumentPath;
                        }
                    }
                }
            }

            // stop if user answered a question with "no"
            if (string.IsNullOrEmpty(wordAdapter.DocumentPath) || (Configuration.ConfigurationTest.SetHyperlinkBase && !wordAdapter.DocumentPath.Equals(wordAdapter.HyperlinkBase)))
            {
                throw new OperationCanceledException(Resources.TestResult_OperationCancelled);
            }
            return isFirstAttachment;
        }

        /// <summary>
        /// Get values for a parametrized query.
        /// This method will look for a specific linktype and a workitem and will return a list of workitems that mathc the case
        /// </summary>
        /// <param name="propertyContainer"></param>
        /// <param name="matches"></param>
        /// <returns></returns>
        private List<WorkItem> GetValuesForParametrizedQuery(object propertyContainer, MatchCollection matches)
        {
            Guard.ThrowOnArgumentNull(propertyContainer, "propertyContainer");
            //Ensure that property Container is of type ITfsTestResultDetail
            if (propertyContainer.GetType() != typeof(TfsTestResultDetail))
            {
                throw new ConfigurationException($"You are trying to use the parameterized property LinkedWorkItemsForTestResult for a WorkItem of type {propertyContainer.GetType()}. This property can only be used for TestResults. Please use the Links property instead");
            }

            var tfsTestResultDetail = propertyContainer as ITfsTestResultDetail;

            var linkType = matches[1].Groups[1].Value;
            var workItemType = matches[0].Groups[1].Value;

            if (tfsTestResultDetail != null)
            {
                var result = TfsTestAdapter.GetAllLinkedWorkItemsForWorkItemId(tfsTestResultDetail.TestResult.TestCaseId, workItemType, linkType, null);
                return result;
            }
            else
            {
                return null;
            }
        }

        #endregion Private replace method

    }
}
