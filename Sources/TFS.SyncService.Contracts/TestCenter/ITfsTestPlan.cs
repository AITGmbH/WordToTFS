using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test suite.
    /// </summary>
    public interface ITfsTestPlan
    {
        /// <summary>
        /// Gets the id of the test plan.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets the name of test plan.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets root test suite.
        /// </summary>
        ITfsTestSuite RootTestSuite { get; }
    }
}
