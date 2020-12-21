using Microsoft.TeamFoundation.TestManagement.Client;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test result.
    /// </summary>
    public interface ITfsTestResult
    {
        /// <summary>
        /// Gets the identifier of test result.
        /// </summary>
        TestCaseResultIdentifier Id { get; }

        int TestCaseId
        {
            get;
        }
    }
}
