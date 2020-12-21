using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of a shared step
    /// </summary>
    public interface ITfsSharedStep
    {
        /// <summary>
        /// Gets the id of the shared steps.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets the title of the shared step.
        /// </summary>
        string Title { get; }

    }
}
