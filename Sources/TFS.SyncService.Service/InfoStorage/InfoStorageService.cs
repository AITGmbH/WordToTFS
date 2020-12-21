using System;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.InfoStorage;

namespace AIT.TFS.SyncService.Service.InfoStorage
{
    /// <summary>
    /// Class implements functionality for information storage service to store information created during a long-lasting process like publishing.
    /// </summary>
    internal class InfoStorageService : IInfoStorageService
    {
        #region fields

        private readonly IInfoCollection<IUserInformation> _infoCollection = new InfoCollection<IUserInformation>();

        #endregion

        /// <summary>
        /// Gets the user information.
        /// </summary>
        public IInfoCollection<IUserInformation> UserInformation
        {
            get
            {
                return _infoCollection;
            }
        }

        /// <summary>
        /// Removes all information items.
        /// </summary>
        public void ClearAll()
        {
            _infoCollection.Clear();
        }

        /// <summary>
        /// Adds an item to the collection.
        /// </summary>
        /// <param name="information">The item to add</param>
        public void AddItem(IUserInformation information)
        {
            _infoCollection.Add(information);
        }

        /// <summary>
        /// Convenience method that adds an exception.
        /// </summary>
        /// <param name="title">Title of the entry.</param>
        /// <param name="ex">Exception to trace.</param>
        public void NotifyError(string title, Exception ex)
        {
            if (ex == null) throw new ArgumentNullException("ex");

            _infoCollection.Add(new UserInformation
                                    {
                                        Type = ex is ConfigurationException ? UserInformationType.Warning : UserInformationType.Error,
                                        Text = title,
                                        Explanation = ex.Message
                                    });
        }

        /// <summary>
        /// Notifies the error.
        /// </summary>
        /// <param name="title">The title.</param>
        /// <param name="ex">The ex.</param>
        /// <param name="navigateAction">The navigate action.</param>
        /// <exception cref="System.NotImplementedException"></exception>
        public void NotifyError(string title, Exception ex, Action navigateAction)
        {
            if (ex == null) throw new ArgumentNullException("ex");

            _infoCollection.Add(new UserInformation
                                    {
                                        Type = ex is ConfigurationException ? UserInformationType.Warning : UserInformationType.Error,
                                        Text = title,
                                        Explanation = ex.Message,
                                        NavigateToSourceAction = navigateAction
                                    });
        }
    }
}