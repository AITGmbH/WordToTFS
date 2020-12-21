namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test plan.
    /// </summary>
    public interface ITfsTestPlanDetail : ITfsPropertyValueProvider
    {
        /// <summary>
        /// Gets the base information
        /// </summary>
        ITfsTestPlan TestPlan { get; }
    }
}
