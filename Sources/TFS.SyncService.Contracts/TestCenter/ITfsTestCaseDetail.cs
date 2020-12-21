using Microsoft.TeamFoundation.TestManagement.Client;
using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test case.
    /// </summary>
    public interface ITfsTestCaseDetail : ITfsPropertyValueProvider
    {
        /// <summary>
        /// Gets the base information.
        /// </summary>
        ITfsTestCase TestCase { get; }

        /// <summary>
        /// Gets the id of the test case.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets the title of test case.
        /// </summary>
        string Title { get; }

        /// <summary>
        /// Gets the iteration path of test case.
        /// </summary>
        string IterationPath { get; }

        /// <summary>
        /// Gets the area path of test case.
        /// </summary>
        string AreaPath { get; }

        /// <summary>
        /// Gets the number of work item.
        /// </summary>
        int WorkItemId { get; }

        /// <summary>
        /// Gets a value indicating whether shared steps will be expanded.
        /// </summary>
        /// <value>
        ///   <c>true</c> if [expand shared steps]; otherwise, <c>false</c>.
        /// </value>
        bool ExpandSharedSteps { get; }

        /// <summary>
        /// Special Property that contains all Test Parameters with all values
        /// 
        /// </summary>

        ITestCaseParameters TestParametersWithAllValues
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the original test case - <see cref="ITestCase"/>.
        /// </summary>
        ITestCase OriginalTestCase { get; }
    }
}
