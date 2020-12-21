using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test suite.
    /// </summary>
    public interface ITfsTestSuite
    {
        /// <summary>
        /// Gets the id of the test suite.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets name of the test suite.
        /// </summary>
        string Title { get; }

        ITfsTestPlan AssociatedTestPlan
        {
            get;
        }
        /// <summary>
        /// Gets all associated test suites.
        /// </summary>
        IEnumerable<ITfsTestSuite> TestSuites { get; }
    }
}
