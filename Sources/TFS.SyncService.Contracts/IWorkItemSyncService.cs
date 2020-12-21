namespace AIT.TFS.SyncService.Contracts
{
    using System.Collections.Generic;
    using System.ServiceModel;
    using Configuration;
    using WorkItemObjects;
    using TestCenter;

    /// <summary>
    /// Interface defines the functionality of the service for synchronization of two adapters.
    /// </summary>
    [ServiceContract]
    public interface IWorkItemSyncService
    {
        /// <summary>
        /// Publish from one adapter to another. This synchronizes the destination adapter, saves it
        /// and then synchronizes the source adapter again. This is, because the destination adapter (team foundation server)
        /// makes automatic changes to work items (revision)
        /// </summary>
        /// <param name="source">
        /// The source of synchronization - synchronize from this work item.
        /// </param>
        /// <param name="destination">
        /// The destination of synchronization - synchronize to this work item.
        /// </param>
        /// <param name="workItems">List of work items to publish</param>
        /// <param name="forceOverwrite">Sets whether work items are synchronized, even if the destination is more recent</param>
        /// <param name="configuration">Configuration to use during publish</param>
        [OperationContract]
        void Publish(IWorkItemSyncAdapter source, IWorkItemSyncAdapter destination, IEnumerable<IWorkItem> workItems, bool forceOverwrite, IConfiguration configuration);

        /// <summary>
        /// Imports work work items from TFS into Word and refreshes existing work items.
        /// </summary>
        /// <param name="sourceTfs">Source adapter represented by TFS server.</param>
        /// <param name="destinationWord">Destination adapter represented by Word document.</param>
        /// <param name="importWorkItems">List of work items to import.</param>
        /// <param name="configuration">Configuration to use during refresh</param>
        void Refresh(IWorkItemSyncAdapter sourceTfs, IWorkItemSyncAdapter destinationWord, IEnumerable<IWorkItem> importWorkItems, IConfiguration configuration);

        /// <summary>
        /// Imports work work items from TFS into Word and refreshes existing work items.
        /// </summary>
        /// <param name="sourceTfs">Source adapter represented by TFS server.</param>
        /// <param name="destinationWord">Destination adapter represented by Word document.</param>
        /// <param name="importWorkItems">List of work items to import.</param>
        /// <param name="configuration">Configuration to use during refresh</param>
        /// <param name="testReportHelper">The test report helper.</param>
        /// <param name="testCases">The test cases.</param>
        void RefreshAndSubstituteTestItems(IWorkItemSyncAdapter sourceTfs, IWorkItemSyncAdapter destinationWord, IEnumerable<IWorkItem> importWorkItems, IConfiguration configuration, ITestReportHelper testReportHelper, IDictionary<int, ITfsTestCaseDetail> testCases);

        /// <summary>
        /// Imports work work items from TFS into Word and refreshes existing work items.
        /// </summary>
        /// <param name="sourceTfs">Source adapter represented by TFS server.</param>
        /// <param name="destinationWord">Destination adapter represented by Word document.</param>
        /// <param name="importWorkItems">List of work items to import.</param>
        /// <param name="configuration">Configuration to use during refresh</param>
        /// <param name="testReportHelper">The test report helper.</param>
        /// <param name="sharedSteps">The shared steps.</param>
        void RefreshAndSubstituteSharedStepItems(IWorkItemSyncAdapter sourceTfs, IWorkItemSyncAdapter destinationWord, IEnumerable<IWorkItem> importWorkItems, IConfiguration configuration, ITestReportHelper testReportHelper, IDictionary<int, ITfsSharedStepDetail> sharedSteps);
    }
}