#region Usings
using System;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;
using System.Linq;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements <see cref="ITfsTestResultDetail"/> - detail information about test result.
    /// </summary>
    public class TfsTestResultDetail : TfsPropertyValueProvider, ITfsTestResultDetail
    {

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestResultDetail"/> class.
        /// </summary>
        /// <param name="testResult">Associated <see cref="TfsTestResult"/>.</param>
        public TfsTestResultDetail(TfsTestResult testResult, int latestTestRunId)
        {
            if (testResult == null)
                throw new ArgumentNullException("testResult");
            TestResultClass = testResult;
            Outcome = TestResultClass.OriginalTestResult.Outcome;
            DateCompleted = TestResultClass.OriginalTestResult.DateCompleted;
            DateCreated = TestResultClass.OriginalTestResult.DateCreated;
            LastUpdated = TestResultClass.OriginalTestResult.LastUpdated;
            TestCase = testResult.OriginalTestResult.GetTestCase();
            TestConfigurationName = testResult.OriginalTestResult.TestConfigurationName;
            TestConfigurationId = testResult.OriginalTestResult.TestConfigurationId;
            LinkedWorkItemsForTestResult = new List<WorkItem>();
            LatestTestRunId = latestTestRunId;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public TfsTestResult TestResultClass { get; private set; }

        /// <summary>
        /// Gets the original test result - <see cref="ITestResult"/>.
        /// </summary>
        public ITestResult OriginalTestResult
        {
            get { return TestResultClass.OriginalTestResult; }
        }

        public ITestCase TestCase
        {
            get;
            private set;

        }

        /// <summary>
        /// Returns the latest test run id of the current test result
        /// </summary>
        public int LatestTestRunId
        {
            get;
            private set;
        }


        #endregion Public properties

        #region Protected override properties

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get { return OriginalTestResult; }
        }

        #endregion Protected override properties

        #region Implementation of ITfsTestCaseDetail

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public ITfsTestResult TestResult
        {
            get { return TestResultClass; }
        }

        /// <summary>
        /// Gets an indication of the outcome of the test.
        /// </summary>
        public TestOutcome Outcome { get; private set; }
        
        /// <summary>
        /// Gets or sets the date the test was completed.
        /// </summary>
        public DateTime DateCompleted { get; private set; }

        /// <summary>
        /// Gets or sets the date the test was created.
        /// </summary>
        public DateTime DateCreated { get; private set; }

        /// <summary>
        /// Gets the date and time that this result was last updated.
        /// </summary>
        public DateTime LastUpdated { get; private set; }

        /// <summary>
        /// The Name of the configuration where the test was ran
        /// </summary>
        public string TestConfigurationName
        {
            get;
            private set;
        }

        /// <summary>
        /// The Name of the configuration where the test was ran
        /// </summary>
        public int TestConfigurationId
        {
            get;
            private set;
        }

        /// <summary>
        /// The linked workitems for this testResult (obtained from the corresponding Test Case)
        /// </summary>
        public List<WorkItem> LinkedWorkItemsForTestResult
        {
            get;
            set;
        }

        /// <summary>
        /// Gets an enumerable of iteration test results
        /// This property should hide ITestIterationResult.Actions and instead expose a list of wrappers for those actions
        /// </summary>
        public IList<TfsTestIterationResult> Iterations
        {
            get
            {
                return TestResultClass.OriginalTestResult.Iterations.Select(x => new TfsTestIterationResult(x, TestCase)).ToList();
            }
        }

        #endregion Implementation of ITfsTestCaseDetail
    }
}
