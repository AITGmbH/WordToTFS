using System;

namespace AIT.TFS.SyncService.Contracts.InfoStorage
{
    /// <summary>
    /// Interface defines functionality for info storage service used to store information for user feedback.
    /// </summary>
    public interface IInfoStorageService
    {
        /// <summary>
        /// Gets all stored information.
        /// </summary>
        IInfoCollection<IUserInformation> UserInformation { get; }

        /// <summary>
        /// Removes all information items.
        /// </summary>
        void ClearAll();

        /// <summary>
        /// Adds an item to the collection.
        /// </summary>
        /// <param name="information">The item to add</param>
        void AddItem(IUserInformation information);

        /// <summary>
        /// Convenience method that adds an exception.
        /// </summary>
        /// <param name="title">Title of the entry.</param>
        /// <param name="ex">Exception to trace.</param>
        void NotifyError(string title, Exception ex);

        /// <summary>
        /// Convenience method that adds an exception.
        /// </summary>
        /// <param name="title">Title of the entry.</param>
        /// <param name="ex">Exception to trace.</param>
        /// <param name="navigateAction">An action to execute when the user clicks the navigateTo-button in the error view</param>
        void NotifyError(string title, Exception ex, Action navigateAction);
    }
}