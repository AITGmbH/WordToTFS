using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test case.
    /// </summary>
    public interface ITfsTestCase
    {
        /// <summary>
        /// Gets the id of the test case.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets the title of the test case.
        /// </summary>
        string Title { get; }

        ITfsTestSuite AssociatedTestSuiteDetail
        {
            get;
        }


    }
}
