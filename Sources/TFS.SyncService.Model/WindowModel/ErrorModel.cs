#region Usings
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows.Threading;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.WindowModelBase;
#endregion
namespace AIT.TFS.SyncService.Model.WindowModel
{
    /// <summary>
    /// ViewModel for errors logged in the InfoStorageService
    /// </summary>
    public class ErrorModel : ExtBaseModel
    {
        #region Private fields
        private readonly Dispatcher _dispatcher;
        #endregion
        #region Constructor
        /// <summary>
        /// Create a new Instance of the ErrorModel
        /// </summary>
        /// <param name="title">Title to be shown on the pane</param>
        /// <param name="dispatcher">The dispatcher of the user interface thread.</param>
        public ErrorModel(string title, Dispatcher dispatcher)
        {
            Title = title;
            _dispatcher = dispatcher;
            Errors = new ObservableCollection<IUserInformation>();

            var storageService = SyncServiceFactory.GetService<IInfoStorageService>();

            // TODO unregister from infostorage when window is hidden
            storageService.UserInformation.CollectionChanged += PublishingInformationOnCollectionChanged;

            _dispatcher.Invoke((Action)(() =>
                                            {
                                                foreach (var item in storageService.UserInformation)
                                                {
                                                    Errors.Add(item);
                                                }
                                            }));

        }
        #endregion
        #region Public properties
        /// <summary>
        /// 
        /// Gets or sets the title of the pane
        /// </summary>
        public string Title { get; private set; }

        /// <summary>
        /// Gets a collection of all errors occurred during the publishing process.
        /// </summary>
        public ObservableCollection<IUserInformation> Errors { get; private set; }
        #endregion
        #region Private methods
        private void PublishingInformationOnCollectionChanged(object sender, NotifyCollectionChangedEventArgs notifyCollectionChangedEventArgs)
        {
            switch (notifyCollectionChangedEventArgs.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    _dispatcher.Invoke((Action)(() => Errors.Add((IUserInformation)notifyCollectionChangedEventArgs.NewItems[0])));
                    break;

                case NotifyCollectionChangedAction.Reset:
                    _dispatcher.Invoke((Action)(() => Errors.Clear()));
                    break;

                case NotifyCollectionChangedAction.Remove:
                    _dispatcher.Invoke((Action)(() =>
                                                    {
                                                        foreach (var delete in notifyCollectionChangedEventArgs.OldItems)
                                                        {
                                                            Errors.Remove((IUserInformation)delete);
                                                        }
                                                    }));
                    break;
            }
        }
        #endregion
    }
}