using System;
using System.Collections.Generic;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test result.
    /// </summary>
    public interface ITfsTestResultDetail : ITfsPropertyValueProvider
    {
        /// <summary>
        /// Gets the base information.
        /// </summary>
        ITfsTestResult TestResult { get; }

        /// <summary>
        /// Gets an indication of the outcome of the test.
        /// </summary>
        TestOutcome Outcome { get; }

        /// <summary>
        /// Gets or sets the date the test was completed.
        /// </summary>
        DateTime DateCompleted { get; }

        /// <summary>
        /// Gets or sets the date the test was created.
        /// </summary>
        DateTime DateCreated { get; }

        /// <summary>
        /// Gets the date and time that this result was last updated.
        /// </summary>
        DateTime LastUpdated { get; }

        /// <summary>
        /// The Name of the configuration where the test was ran
        /// </summary>
        string TestConfigurationName
        {
            get;
        }


        /// <summary>
        /// The Name of the configuration where the test was ran
        /// </summary>
        int TestConfigurationId
        {
            get;
        }

        /// <summary>
        /// Special property that contains all workitems that are linked to this testresult.
        /// The list is obtained by the corresponding testcase of the testresult
        /// </summary>
        List<WorkItem> LinkedWorkItemsForTestResult
        {
            get;
            set;
        }

        /// <summary>
        /// Returns the latest test run id of the current test result
        /// </summary>
        int LatestTestRunId
        {
            get;
        }

    }
}
