#region Usings
using System.ComponentModel;
using System.Threading;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Contracts.Word;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Properties;
using AIT.TFS.SyncService.Model.WindowModelBase;
using AIT.TFS.SyncService.Service.Configuration;
using Microsoft.Office.Interop.Word;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using MessageBox = System.Windows.MessageBox;
using Task = System.Threading.Tasks.Task;
#endregion
namespace AIT.TFS.SyncService.Model.WindowModel
{
    /// <summary>
    /// ViewModel class for GetWorkItemsPanelView view.
    /// </summary>
    public class WorkItemOverviewPanelViewModel : ExtBaseModel
    {
        #region Private fields

        private readonly IQueryConfiguration _queryConfiguration;
        private bool _useLinkedWorkItems;
        private bool _isAnyLinkType = true;
        private bool _isDirectLinkOnlyMode;
        private QueryItem _selectedQuery;
        private WorkItemType _selectedWorkItemType;
        private IList<DataItemModel<SynchronizedWorkItemViewModel>> _foundWorkItems;
        private bool _isFinding;
        private CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();

        #endregion Private fields
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="GetWorkItemsPanelViewModel"/> class.
        /// </summary>
        /// <param name="tfsService">Team foundation service.</param>
        /// <param name="documentModel">Model of word document to work with.</param>
        /// <param name="wordRibbon">Word ribbon.</param>
        public WorkItemOverviewPanelViewModel(ITfsService tfsService, ISyncServiceDocumentModel documentModel, IWordRibbon wordRibbon)
        {
            if (tfsService == null)
                throw new ArgumentNullException("tfsService");
            if (documentModel == null)
                throw new ArgumentNullException("documentModel");
            if (wordRibbon == null)
                throw new ArgumentNullException("wordRibbon");

            TfsService = tfsService;
            WordRibbon = wordRibbon;

            WordDocument = documentModel.WordDocument as Document;
            DocumentModel = documentModel;
            DocumentModel.PropertyChanged += DocumentModelOnPropertyChanged;

            _queryConfiguration = new QueryConfiguration
                                      {
                                          UseLinkedWorkItems = false,
                                          IsDirectLinkOnlyMode = false
                                      };
            DocumentModel.ReadQueryConfiguration(_queryConfiguration);
            SelectedQuery = FindQueryInHierarchy(QueryHierarchy, _queryConfiguration.QueryPath);

            LinkTypes = TfsService.GetLinkTypes().Select(
                x => new DataItemModel<ITFSWorkItemLinkType>(x)
                         {
                             IsChecked = QueryConfiguration.LinkTypes.Contains(x.ReferenceName)
                         }).ToList();
        }

        /// <summary>
        /// Abort sync state queries if user starts an actual synchronizing operation
        /// </summary>
        private void DocumentModelOnPropertyChanged(object sender, PropertyChangedEventArgs propertyChangedEventArgs)
        {
            if (propertyChangedEventArgs.PropertyName.Equals("OperationInProgress"))
            {
                if (DocumentModel.OperationInProgress)
                {
                    _cancellationTokenSource.Cancel();
                }

                OnPropertyChanged(nameof(CanFindWorkItems));
            }
        }

        /// <summary>
        /// Recursively searches for the query in the query hierarchy that has the given path
        /// </summary>
        private QueryDefinition FindQueryInHierarchy(IEnumerable<QueryItem> queryFolder, string queryPath)
        {
            if (queryFolder == null) return null;
            foreach (var item in queryFolder)
            {
                if (item is QueryDefinition && item.Path == queryPath) return (QueryDefinition)item;

                if (item is QueryFolder)
                {
                    var foundItem = FindQueryInHierarchy(item as QueryFolder, queryPath);
                    if (foundItem != null) return foundItem;
                }
            }
            return null;
        }

        #endregion Constructors
        #region Public properties

        /// <summary>
        /// Gets or sets Word ribbon service.
        /// </summary>
        public IWordRibbon WordRibbon { get; private set; }

        /// <summary>
        /// Gets or sets team foundation server service.
        /// </summary>
        public ITfsService TfsService { get; private set; }

        /// <summary>
        /// Gets model of associated word document.
        /// </summary>
        public ISyncServiceDocumentModel DocumentModel { get; private set; }

        /// <summary>
        /// Gets the associated word document.
        /// </summary>
        public Document WordDocument { get; private set; }

        /// <summary>
        /// Gets query configuration object.
        /// </summary>
        public IQueryConfiguration QueryConfiguration
        {
            get
            {
                return _queryConfiguration;
            }
        }

        /// <summary>
        /// Gets all query hierarchy.
        /// </summary>
        public QueryHierarchy QueryHierarchy
        {
            get
            {
                return TfsService.GetWorkItemQueries();
            }
        }

        /// <summary>
        /// Gets all link types.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "You are not supposed to manipulate the list.")]
        public IList<DataItemModel<ITFSWorkItemLinkType>> LinkTypes
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets whether to use linked work items.
        /// </summary>
        public bool UseLinkedWorkItems
        {
            get
            {
                return _useLinkedWorkItems;
            }
            set
            {
                _useLinkedWorkItems = value;
                OnPropertyChanged(nameof(UseLinkedWorkItems));
                OnPropertyChanged(nameof(IsCustomLinksListEnabled));
            }
        }

        /// <summary>
        /// Gets or sets whether is set any link type.
        /// </summary>
        public bool IsAnyLinkType
        {
            get
            {
                return _isAnyLinkType;
            }
            set
            {
                _isAnyLinkType = value;
                OnPropertyChanged(nameof(IsAnyLinkType));
                OnPropertyChanged(nameof(IsCustomLinkType));
                OnPropertyChanged(nameof(IsCustomLinksListEnabled));
            }
        }

        /// <summary>
        /// Gets or sets whether to get only direct links only or recursive
        /// </summary>
        public bool IsDirectLinkOnlyMode
        {
            get
            {
                return _isDirectLinkOnlyMode;
            }
            set
            {
                _isDirectLinkOnlyMode = value;
                OnPropertyChanged(nameof(IsDirectLinkOnlyMode));
            }
        }

        /// <summary>
        /// Gets whether is set custom link type.
        /// </summary>
        public bool IsCustomLinkType
        {
            get
            {
                return !_isAnyLinkType;
            }
        }

        /// <summary>
        /// Gets whether custom links list control is enabled.
        /// </summary>
        public bool IsCustomLinksListEnabled
        {
            get
            {
                return IsCustomLinkType && UseLinkedWorkItems;
            }
        }

        /// <summary>
        /// Gets or sets actual selected query.
        /// </summary>
        public QueryItem SelectedQuery
        {
            get
            {
                return _selectedQuery;
            }
            set
            {
                _selectedQuery = value;
                OnPropertyChanged(nameof(SelectedQuery));
                OnPropertyChanged(nameof(CanFindWorkItems));
            }
        }

        /// <summary>
        /// Gets whether user can start find work items based on criteria.
        /// </summary>
        public bool CanFindWorkItems
        {
            get
            {
                return !IsFinding && !DocumentModel.OperationInProgress;
            }
        }

        /// <summary>
        /// Gets all work item types.
        /// </summary>
        public WorkItemTypeCollection AllWorkItemTypes
        {
            get
            {
                return TfsService.GetAllWorkItemTypes();
            }
        }

        /// <summary>
        /// Gets or sets selected work item type.
        /// </summary>
        public WorkItemType SelectedWorkItemType
        {
            get
            {
                return _selectedWorkItemType;
            }
            set
            {
                _selectedWorkItemType = value;
                OnPropertyChanged(nameof(SelectedWorkItemType));
                OnPropertyChanged(nameof(CanFindWorkItems));
            }
        }

        /// <summary>
        /// Gets collection of all found WorkItems.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "You are not supposed to manipulate the list.")]
        public IList<DataItemModel<SynchronizedWorkItemViewModel>> FoundWorkItems
        {
            get
            {
                return _foundWorkItems;
            }
            private set
            {
                _foundWorkItems = value;
                OnPropertyChanged(nameof(FoundWorkItems));
            }
        }

        /// <summary>
        /// Gets or sets whether is Find process running.
        /// </summary>
        public bool IsFinding
        {
            get
            {
                return _isFinding;
            }
            set
            {
                _isFinding = value;
                OnPropertyChanged(nameof(IsFinding));
                OnPropertyChanged(nameof(CanFindWorkItems));
            }
        }

        #endregion Public properties
        #region Public methods


        /// <summary>
        /// Save data to query configuration.
        /// </summary>
        private void SaveQueryConfiguration()
        {
            QueryConfiguration.QueryPath = SelectedQuery != null ? SelectedQuery.Path : null;
            QueryConfiguration.UseLinkedWorkItems = UseLinkedWorkItems;
            QueryConfiguration.IsDirectLinkOnlyMode = IsDirectLinkOnlyMode;
            QueryConfiguration.LinkTypes.Clear();
            if (QueryConfiguration.UseLinkedWorkItems && IsCustomLinkType)
            {
              foreach (var link in LinkTypes)
              {
                 if (link.IsChecked)
                 {
                     QueryConfiguration.LinkTypes.Add(link.Item.ReferenceName);
                 }
              }
            }
            QueryConfiguration.ImportOption = QueryImportOption.SavedQuery;
            QueryConfiguration.WorkItemType = SelectedWorkItemType != null ? SelectedWorkItemType.Name : null;
        }

        /// <summary>
        /// Execute find work items process to fill up the Find list box.
        /// </summary>
        public void FindWorkItems()
        {
            SaveQueryConfiguration();
            SyncServiceTrace.I(Resources.FindWorkItems);

            if (SelectedQuery == null)
            {
                QueryConfiguration.ImportOption = QueryImportOption.IDs;

                var wordAdapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(DocumentModel.WordDocument, DocumentModel.Configuration);
                wordAdapter.Open(null);

                QueryConfiguration.ByIDs.Clear();

                foreach (var workItem in wordAdapter.WorkItems)
                {
                    if (!workItem.IsNew)
                    {
                        QueryConfiguration.ByIDs.Add(workItem.Id);
                    }
                }
            }

            // check whether is at least one link type selected
            if (QueryConfiguration.UseLinkedWorkItems && IsCustomLinkType && QueryConfiguration.LinkTypes.Count == 0)
            {
                MessageBox.Show(Resources.WordToTFS_Error_LinkTypesNotSelected, Resources.MessageBox_Title, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // disable all other functionality
            if (WordRibbon != null)
            {
                WordRibbon.ResetBeforeOperation(DocumentModel);
                WordRibbon.DisableAllControls(DocumentModel);
            }

            // display progress dialog
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.ShowProgress();
            progressService.NewProgress(Resources.QueryWorkItems_Title);
            IsFinding = true;

            Task.Factory.StartNew(DoFind).ContinueWith(FindFinished, TaskScheduler.FromCurrentSynchronizationContext());
        }

        /// <summary>
        /// Cancels all query operations
        /// </summary>
        public void Cancel()
        {
            _cancellationTokenSource.Cancel();
        }

        #endregion Public methods
        #region Private event handler methods

        /// <summary>
        /// Background method to find the work items from TFS.
        /// </summary>
        private void DoFind()
        {
            SyncServiceTrace.D(Resources.DoFindProccessMessage);
            DocumentModel.SaveQueryConfiguration(_queryConfiguration);

            //var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(WordDocument);
            var config = DocumentModel.Configuration;
            var tfsService = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(DocumentModel.TfsServer, DocumentModel.TfsProject, _queryConfiguration, config);
            var wordAdapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(DocumentModel.WordDocument, DocumentModel.Configuration);
            wordAdapter.Open(null);
            tfsService.Open(null);

            EnsureTfsLazyLoadingFinished(tfsService);

            // Cancel all queries in case there are still some running
            Cancel();
            _cancellationTokenSource = new CancellationTokenSource();
            var token = _cancellationTokenSource.Token;

            var tfsItems = (from wi in tfsService.WorkItems
                            select new DataItemModel<SynchronizedWorkItemViewModel>(new SynchronizedWorkItemViewModel(wi, null, wordAdapter, DocumentModel, token))
                                       {
                                           IsChecked = true
                                       }).ToList();

            var wordItems = (from wi in wordAdapter.WorkItems
                             where wi.IsNew
                             select new DataItemModel<SynchronizedWorkItemViewModel>(new SynchronizedWorkItemViewModel(null, wi, wordAdapter, DocumentModel, token))
                                        {
                                            IsChecked = true
                                        }).ToList();

            FoundWorkItems = tfsItems.Concat(wordItems).ToList();
        }

        /// <summary>
        /// This beautiful piece of code ensures that the TFS API has finished creating its field collection.
        /// Not doing this sequentially before creating parallel synchronization tasks will result in a
        /// "Collection has been modified" exception.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "x")]
        private static void EnsureTfsLazyLoadingFinished(IWorkItemSyncAdapter tfsService)
        {
            if (tfsService.WorkItems.Any())
            {
                foreach (var field in tfsService.WorkItems.First().Fields)
                {
// ReSharper disable UnusedVariable
                    var x = field.Value;
// ReSharper restore UnusedVariable
                }
            }
        }

        /// <summary>
        /// Occurs when the background operation has completed, has been canceled, or has raised an exception.
        /// </summary>
        private void FindFinished(Task task)
        {
            IsFinding = false;
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.HideProgress();

            // enable all other functionality
            if (WordRibbon != null) WordRibbon.EnableAllControls(DocumentModel);

            SyncServiceTrace.I(Resources.FindFinished);
        }

        #endregion Private event handler methods
    }
}