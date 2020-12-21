using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.TfsHelper;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using Document = Microsoft.Office.Interop.Word.Document;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for PublishingStateControl.
    /// </summary>
    public partial class PublishingStateControl : IHostPaneControl
    {
        #region Fields
        private readonly IInfoStorageService _storageService;
        private readonly ObservableCollection<IUserInformation> _infos = new ObservableCollection<IUserInformation>();
        private delegate void AddItemDelegate(IUserInformation information);
        private bool _isInProgress = false;
        #endregion Fields

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="PublishingStateControl"/> class.
        /// </summary>
        /// <param name="document">The document this panel is attached to.</param>
        public PublishingStateControl(Document document)
        {
            InitializeComponent();
            
            AttachedDocument = document;
            var syncService = SyncServiceFactory.GetService<ISyncServiceModel>();
            DocumentModel = syncService.GetModel(document);

            _storageService = SyncServiceFactory.GetService<IInfoStorageService>();
            _storageService.ClearAll();
            DataContext = this;
        }
        #endregion Constructor

        #region Public properties

        /// <summary>
        /// Gets the model of associated word document.
        /// </summary>
        public ISyncServiceDocumentModel DocumentModel { get; private set; }

        /// <summary>
        /// Gets the publish information entries
        /// </summary>
        public ObservableCollection<IUserInformation> PublishInformation
        {
            get 
            {
                return _infos; 
            }
        }
        /// <summary>
        /// Sets a value indicating whether this instance is in progress.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is in progress; otherwise, <c>false</c>.
        /// </value>
        public bool IsInProgress
        {
            get { return _isInProgress; }
            set
            {
                if (value)
                {
                    // when publishing starts then register event
                    if (_storageService != null && _storageService.UserInformation != null)
                        _storageService.UserInformation.CollectionChanged += OnCollectionChanged;
                }
                else
                {
                    // when publishing ends then unregister event
                    if (_storageService != null && _storageService.UserInformation != null)
                        _storageService.UserInformation.CollectionChanged -= OnCollectionChanged;
                }

                _isInProgress = value;
                IsEnabled = !value;
            }
        }
        #endregion Public properties

        #region Private methods
        /// <summary>
        /// Adds the item.
        /// </summary>
        /// <param name="information">The information.</param>
        private void AddItem(IUserInformation information)
        {
            if ((Dispatcher == null) || Dispatcher.Thread.Equals(System.Threading.Thread.CurrentThread))
            {
                _infos.Add(information);
            }
            else
            {
                Dispatcher.Invoke(new AddItemDelegate(AddItem), information);
            }
        }
        /// <summary>
        /// Predicate determining if an items is filtered from a list.
        /// </summary>
        /// <param name="item">Item to check</param>
        /// <returns>False, if the item is to be filtered out. True otherwise</returns>
        private bool FilterItem(object item)
        {
            switch(((IUserInformation)item).Type)
            {
                case UserInformationType.Unmodified:
                    return FilterSkippedCheckBox.IsChecked == true;

                case UserInformationType.Success:
                    return FilterSuccessCheckBox.IsChecked == true;

                case UserInformationType.Error:
                    return FilterFailedCheckBox.IsChecked == true;

                case UserInformationType.Conflicting:
                    return FilterConflictingCheckBox.IsChecked == true;
            }

            return true;
        }
        #endregion Private methods

        #region Event handlers
        /// <summary>
        /// Filters the list view.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void FilterListView(object sender, RoutedEventArgs e)
        {
            if (PublishedWorkItemsListView == null)
                return;
            ICollectionView listCollectionView = CollectionViewSource.GetDefaultView(PublishedWorkItemsListView.ItemsSource);
            listCollectionView.Filter = FilterItem;
        }

        private void CanExecuteFind(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = e.Parameter is IUserInformation;
        }

        private void ExecutedFind(object sender, ExecutedRoutedEventArgs e)
        {
            var pubinfo = e.Parameter as IUserInformation;
            if ((pubinfo == null) || (pubinfo.Source == null) || (AttachedDocument == null)) return;

            // if is defined Field then jump directly to it
            if (pubinfo.NavigateToSourceAction != null)
            {
                pubinfo.NavigateToSourceAction.Invoke();
                return;
            }

            //if there is a bookmark with name of the workitem title
            if (!string.IsNullOrEmpty(pubinfo.Source.Id.ToString(CultureInfo.InvariantCulture)) && pubinfo.Source.Id != 0)
            {
                object title = string.Concat("w2t", pubinfo.Source.Id.ToString(CultureInfo.InvariantCulture));

                if (AttachedDocument.Bookmarks.Exists(title.ToString()))
                {
                    AttachedDocument.Bookmarks[title].Select();
                }
                else
                {
                    //if there is no bookmark, but it has an Id -> Find ID!
                    Microsoft.Office.Interop.Word.Range range = AttachedDocument.Range();
                    var find = range.Find;
                    find.Text = pubinfo.Source.Id.ToString(CultureInfo.InvariantCulture);
                    if (find.Execute())
                    {
                        range.Select();
                    }
                }
            }
            else
            {
                //If not propper synced, search for the Title
                Microsoft.Office.Interop.Word.Range range = AttachedDocument.Range();
                var find = range.Find;
                if (!pubinfo.Source.Fields.Contains(FieldReferenceNames.SystemTitle))
                    return;
                find.Text = pubinfo.Source.Fields[FieldReferenceNames.SystemTitle].Value;
                if (find.Execute())
                {
                    range.Select();
                }
            }
        }
        /// <summary>
        /// Called when Unloaded event fired.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void UserControlUnloaded(object sender, RoutedEventArgs e)
        {
            // unregister event delegate to ensure that this obejct will be garbage collected
            if (_storageService != null && _storageService.UserInformation != null)
                _storageService.UserInformation.CollectionChanged -= OnCollectionChanged;
        }
        /// <summary>
        /// Called when publishing information collection changed.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="System.Collections.Specialized.NotifyCollectionChangedEventArgs"/> instance containing the event data.</param>
        private void OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e != null && e.NewItems != null)
                AddItem(e.NewItems[0] as IUserInformation);
        }

        /// <summary>
        /// Open browser and navigate to a page that allows to edit the destination work item of publish information.
        /// </summary>
        private void ExecuteBrowseHome(object sender, ExecutedRoutedEventArgs e)
        {
            var pubinfo = e.Parameter as IUserInformation;
            if ((pubinfo == null) || (pubinfo.Destination == null)) return;

            // Get link to team foundation server web access and insert work item id into url.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(AttachedDocument);
            var source = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(DocumentModel.TfsServer, DocumentModel.TfsProject, null, config);
            var url = source.GetWorkItemEditorUrl(pubinfo.Destination.Id);
            System.Diagnostics.Process.Start(url.ToString());
        }

        #endregion Event handlers

        #region Implementation of IHostPaneControl interface
        
        /// <summary>
        /// Gets or sets Word Document instance.
        /// </summary>
        public Document AttachedDocument { get; private set; }

        /// <summary>
        /// The method is called after the visibility of the containing panel has changed.
        /// </summary>
        /// <param name="isVisible"><c>true</c>if the panel containing this view is shown. False otherwise.</param>
        public void VisibilityChanged(bool isVisible)
        {
            // we don't care
        }

        #endregion Implementation of IHostPaneControl interface



    }
}