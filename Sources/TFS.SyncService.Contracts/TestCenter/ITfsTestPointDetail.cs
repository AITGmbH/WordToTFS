
namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test case.
    /// </summary>
    public interface ITfsTestPointDetail : ITfsPropertyValueProvider
    {
        /// <summary>
        /// Gets the base information.
        /// </summary>
        ITfsTestPoint TestPoint { get; }

        /// <summary>
        /// Gets the id of the test case.
        /// </summary>
        int Id { get; }

    }
}
