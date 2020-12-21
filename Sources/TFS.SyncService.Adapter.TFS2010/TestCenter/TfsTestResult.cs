using System;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements <see cref="ITfsTestResult"/>.
    /// </summary>
    public class TfsTestResult : ITfsTestResult
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestResult"/> class.
        /// </summary>
        /// <param name="testResult">Original test result - <see cref="ITestResult"/>.</param>
        /// <param name="testCaseId">The id of the corresponding test case</param>
        public TfsTestResult(ITestCaseResult testResult,int testCaseId)
        {
            if (testResult == null)
                throw new ArgumentNullException("testResult");
            OriginalTestResult = testResult;
            Id = new TestCaseResultIdentifier(OriginalTestResult.TestRunId, OriginalTestResult.TestResultId);
            TestCaseId = testCaseId;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the original test result - <see cref="ITestResult"/>.
        /// </summary>
        public ITestCaseResult OriginalTestResult { get; private set; }

        #endregion Public properties

        #region Implementation of ITfsTestResult

        /// <summary>
        /// Gets the identifier of test result.
        /// </summary>
        public TestCaseResultIdentifier Id { get; private set; }

        public int TestCaseId
        {
            get;
            private set;
        }

        #endregion Implementation of ITfsTestResult
    }
}
