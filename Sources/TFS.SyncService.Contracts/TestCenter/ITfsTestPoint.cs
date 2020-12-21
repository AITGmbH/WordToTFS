
namespace AIT.TFS.SyncService.Contracts.TestCenter
{

    /// <summary>
    /// A single test point that is responsible for the linking of test results to test cases
    /// </summary>
    public interface ITfsTestPoint
    {
        /// <summary>
        /// Gets the id of test point
        /// </summary>
        int Id { get; }



    }
}
