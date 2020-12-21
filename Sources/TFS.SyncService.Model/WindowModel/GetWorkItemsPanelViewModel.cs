#region Usings
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using AIT.TFS.SyncService.Adapter.Word2007;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Contracts.Word;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Model.Properties;
using AIT.TFS.SyncService.Model.WindowModelBase;
using AIT.TFS.SyncService.Service.Configuration;
using Microsoft.Office.Interop.Word;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using MessageBox = System.Windows.MessageBox;
using Task = System.Threading.Tasks.Task;
using AIT.TFS.SyncService.Contracts.Adapter;
using AIT.TFS.SyncService.Contracts.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.Enums;
#endregion
namespace AIT.TFS.SyncService.Model.WindowModel
{
    /// <summary>
    /// ViewModel class for GetWorkItemsPanelView view.
    /// </summary>
    public class GetWorkItemsPanelViewModel : ExtBaseModel
    {
        #region Private fields

        private readonly IQueryConfiguration _queryConfiguration;
        private bool _useLinkedWorkItems;
        private bool _isDirectLinkOnlyMode;
        private bool _isAnyLinkType = true;
        private QueryItem _selectedQuery;
        private bool _isImporting;
        private QueryImportOption _importOption;
        private string _importIDs;
        private string _importTitleContains;
        private WorkItemType _selectedWorkItemType;
        private IList<DataItemModel<SynchronizedWorkItemViewModel>> _foundWorkItems;
        private bool _isFinding;
        private CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();
        private IWordSyncAdapter _importWordAdapter;
        private SynchronizationState _synchronizationState;

        #endregion Private fields
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="GetWorkItemsPanelViewModel"/> class.
        /// </summary>
        /// <param name="tfsService">Team foundation service.</param>
        /// <param name="documentModel">Model of word document to work with.</param>
        /// <param name="wordRibbon">Word ribbon.</param>
        public GetWorkItemsPanelViewModel(ITfsService tfsService, ISyncServiceDocumentModel documentModel, IWordRibbon wordRibbon)
        {
            if (tfsService == null) throw new ArgumentNullException("tfsService");
            if (documentModel == null) throw new ArgumentNullException("documentModel");
            if (wordRibbon == null) throw new ArgumentNullException("wordRibbon");
            TfsService = tfsService;
            WordRibbon = wordRibbon;

            WordDocument = documentModel.WordDocument as Document;
            DocumentModel = documentModel;
            _isDirectLinkOnlyMode = DocumentModel.Configuration.GetDirectLinksOnly;
            DocumentModel.PropertyChanged += DocumentModelOnPropertyChanged;

            _queryConfiguration = new QueryConfiguration
            {
                UseLinkedWorkItems = false
            };

            DocumentModel.ReadQueryConfiguration(_queryConfiguration);
            SelectedQuery = FindQueryInHierarchy(QueryHierarchy, _queryConfiguration.QueryPath);

            var resultOfGetLinkTypes = TfsService.GetLinkTypes();
            if (resultOfGetLinkTypes != null)
            {
                LinkTypes = resultOfGetLinkTypes.Select(
                x => new DataItemModel<ITFSWorkItemLinkType>(x)
                {
                    IsChecked = QueryConfiguration.LinkTypes.Contains(x.ReferenceName)
                }).ToList();
            }

            CollapseQueryTree = DocumentModel.Configuration.CollapsQueryTree;
            IsExistNotShownWorkItems = false;
            SynchronizationState = SynchronizationState.Unknown;
        }

        #endregion Constructors
        #region Public properties

        /// <summary>
        /// Gets or sets the amount of not shown work items.
        /// </summary>
        public string AmountOfWorkItems { get; set; }

        /// <summary>
        /// Gets or sets whether is exist work items that are not shown.
        /// </summary>
        public bool IsExistNotShownWorkItems { get; set; }

        /// <summary>
        /// Gets or sets the synchronization state of this work item.
        /// </summary>
        public SynchronizationState SynchronizationState
        {
            get
            {
                return _synchronizationState;
            }
            set
            {
                _synchronizationState = value;
                TriggerPropertyChanged(nameof(SynchronizationState));
            }
        }

        /// <summary>
        /// Gets or sets the result of comparing work items.
        /// </summary>
        public string ResultOfComparing { get; set; }

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
        /// Property used to show query tree collapsed or not expanded
        /// </summary>
        public bool CollapseQueryTree { get; set; }

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
                OnPropertyChanged(nameof(CanImport));
                OnPropertyChanged(nameof(CanFindWorkItems));
            }
        }

        /// <summary>
        /// Gets whether is possible to import.
        /// </summary>
        public bool CanImport
        {
            get
            {
                return FoundWorkItems != null &&
                       FoundWorkItems.Count > 0 &&
                       !IsImporting &&
                       !IsFinding &&
                       DocumentModel.OperationInProgress == false &&
                       FoundWorkItems.Any(item => item.IsChecked);
            }
        }

        /// <summary>
        /// Gets whether user can start find work items based on criteria.
        /// </summary>
        public bool CanFindWorkItems
        {
            get
            {
                if (IsImporting || IsFinding || DocumentModel.OperationInProgress) return false;

                switch (ImportOption)
                {
                    case QueryImportOption.SavedQuery:
                        var queryDefinition = SelectedQuery as QueryDefinition;
                        if (queryDefinition != null)
                        {
                            return true;
                        }
                        break;
                    case QueryImportOption.IDs:
                        if (string.IsNullOrEmpty(ImportIDs) == false) return true;
                        break;
                    case QueryImportOption.TitleContains:
                        if (string.IsNullOrEmpty(ImportTitleContains) == false) return true;
                        break;
                }

                return false;
            }
        }

        /// <summary>
        /// Gets whether can select found work items.
        /// </summary>
        public bool CanSelectFoundWorkItems
        {
            get
            {
                return FoundWorkItems != null &&
                       FoundWorkItems.Count > 0 &&
                       !IsImporting &&
                       !IsFinding &&
                       FoundWorkItems.Any(item => item.IsChecked == false);
            }
        }

        /// <summary>
        /// Gets whether can unselect found work items.
        /// </summary>
        public bool CanUnselectFoundWorkItems
        {
            get
            {
                return FoundWorkItems != null &&
                       FoundWorkItems.Count > 0 &&
                       !IsImporting &&
                       !IsFinding &&
                       FoundWorkItems.Any(item => item.IsChecked);
            }
        }

        /// <summary>
        /// Gets or sets whether is importing.
        /// </summary>
        public bool IsImporting
        {
            get
            {
                return _isImporting;
            }
            set
            {
                _isImporting = value;
                OnPropertyChanged(nameof(IsImporting));
                OnPropertyChanged(nameof(CanImport));
                OnPropertyChanged(nameof(CanFindWorkItems));
                OnPropertyChanged(nameof(CanSelectFoundWorkItems));
                OnPropertyChanged(nameof(CanUnselectFoundWorkItems));
            }
        }

        /// <summary>
        /// Gets or sets import option.
        /// </summary>
        public QueryImportOption ImportOption
        {
            get
            {
                return _importOption;
            }
            set
            {
                if (value != QueryImportOption.IDs)
                {
                    ImportIDs = string.Empty;
                }

                if (value != QueryImportOption.TitleContains)
                {
                    ImportTitleContains = string.Empty;
                }

                _importOption = value;
                OnPropertyChanged(nameof(ImportOption));
                OnPropertyChanged(nameof(CanFindWorkItems));
            }
        }

        /// <summary>
        /// Gets or sets import ids string.
        /// </summary>
        public string ImportIDs
        {
            get
            {
                return _importIDs;
            }
            set
            {
                _importIDs = value;
                OnPropertyChanged(nameof(ImportIDs));
                OnPropertyChanged(nameof(CanFindWorkItems));
            }
        }

        /// <summary>
        /// Gets or sets import title contains string.
        /// </summary>
        public string ImportTitleContains
        {
            get
            {
                return _importTitleContains;
            }
            set
            {
                _importTitleContains = value;
                OnPropertyChanged(nameof(ImportTitleContains));
                OnPropertyChanged(nameof(CanFindWorkItems));
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
                OnPropertyChanged(nameof(CanImport));
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
                OnPropertyChanged(nameof(CanImport));
                OnPropertyChanged(nameof(CanFindWorkItems));
                OnPropertyChanged(nameof(CanSelectFoundWorkItems));
                OnPropertyChanged(nameof(CanUnselectFoundWorkItems));
            }
        }

        #endregion Public properties
        #region Public methods

        /// <summary>
        /// The method calls some OnPropertyChanged methods.
        /// </summary>
        public void UpdateFromFoundWorkItemsControl()
        {
            OnPropertyChanged(nameof(CanImport));
            OnPropertyChanged(nameof(CanSelectFoundWorkItems));
            OnPropertyChanged(nameof(CanUnselectFoundWorkItems));
        }

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
            QueryConfiguration.ImportOption = ImportOption;
            QueryConfiguration.WorkItemType = SelectedWorkItemType != null ? SelectedWorkItemType.Name : null;
        }

        /// <summary>
        /// Execute find work items process to fill up the Find list box.
        /// </summary>
        public void FindWorkItems()
        {
            SaveQueryConfiguration();
            SyncServiceTrace.I(Resources.FindWorkItems);

            // check whether is at least one link type selected
            if (QueryConfiguration.UseLinkedWorkItems && IsCustomLinkType && QueryConfiguration.LinkTypes.Count == 0)
            {
                MessageBox.Show(Resources.WordToTFS_Error_LinkTypesNotSelected, Resources.MessageBox_Title, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // parse IDs
            if (ImportOption == QueryImportOption.IDs && !ParseIDs(QueryConfiguration.ByIDs)) return;

            if (ImportOption == QueryImportOption.TitleContains)
            {
                QueryConfiguration.ByTitle = ImportTitleContains;
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
            progressService.NewProgress(Resources.GetWorkItems_Title);
            IsFinding = true;

            TaskScheduler scheduler = null;
            try
            {
                scheduler = TaskScheduler.FromCurrentSynchronizationContext();
            }
            catch (InvalidOperationException)
            {
                scheduler = TaskScheduler.Current;
            }

            Task.Factory.StartNew(DoFind).ContinueWith(FindFinished, scheduler);
        }

        /// <summary>
        /// Execute import work items from TFS to Word process.
        /// </summary>
        public void Import()
        {
            SyncServiceTrace.I(Resources.ImportWorkItems);
            SyncServiceFactory.GetService<IInfoStorageService>().ClearAll();

            // check whether cursor is not inside the table
            if (WordSyncHelper.IsCursorInTable(WordDocument.ActiveWindow.Selection))
            {
                SyncServiceTrace.W(Resources.WordToTFS_Error_TableInsertInTable);
                MessageBox.Show(Resources.WordToTFS_Error_TableInsertInTable, Resources.MessageBox_Title, MessageBoxButton.OK, MessageBoxImage.Error);
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
            progressService.NewProgress(Resources.GetWorkItems_Title);
            IsImporting = true;

            // start background thread to execute the import
            var backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += DoImport;
            backgroundWorker.RunWorkerCompleted += ImportFinished;
            backgroundWorker.RunWorkerAsync(backgroundWorker);
        }

        /// <summary>
        /// Creates hyperlinks to selected work items at the current cursor position.
        /// </summary>
        public void CreateHyperlink()
        {
            SyncServiceTrace.I(Resources.CreateHyperlinks);

            // set actual selected configuration
            var ignoreFormatting = DocumentModel.Configuration.IgnoreFormatting;
            var conflictOverwrite = DocumentModel.Configuration.ConflictOverwrite;
            DocumentModel.Configuration.ActivateMapping(DocumentModel.MappingShowName);
            DocumentModel.Configuration.IgnoreFormatting = ignoreFormatting;
            DocumentModel.Configuration.ConflictOverwrite = conflictOverwrite;
            DocumentModel.Configuration.IgnoreFormatting = ignoreFormatting;

            var source = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(DocumentModel.TfsServer, DocumentModel.TfsProject, null, DocumentModel.Configuration);
            var destination = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(WordDocument, DocumentModel.Configuration);
            var importWorkItems = (from wim in FoundWorkItems
                                   where wim.IsChecked
                                   select wim.Item).ToList();


            // search for already existing work items in word and ask whether to overide them
            if (!(source.Open(importWorkItems.Select(x => x.TfsItem.Id).ToArray()) && destination.Open(null))) return;

            foreach (var workItem in importWorkItems)
            {
                destination.CreateWorkItemHyperlink(workItem.TfsItem.Title, source.GetWorkItemEditorUrl(workItem.TfsItem.Id));
            }
        }

        /// <summary>
        /// Cancel all work item synchronization queries.
        /// </summary>
        public void Cancel()
        {
            _cancellationTokenSource.Cancel();
        }

        #endregion Public methods
        #region Private methods

        /// <summary>
        /// Recursively searches for the query in the query hierarchy that has the given path
        /// </summary>
        private QueryDefinition FindQueryInHierarchy(IEnumerable<QueryItem> queryFolder, string queryPath)
        {
            if (queryFolder == null) return null;
            foreach (var item in queryFolder)
            {
                if (item is QueryDefinition && item.Path == queryPath)
                {
                    return (QueryDefinition)item;
                }

                if (item is QueryFolder)
                {
                    var foundItem = FindQueryInHierarchy(item as QueryFolder, queryPath);
                    if (foundItem != null)
                    {
                        return foundItem;
                    }
                }
            }
            return null;
        }

        private bool ParseIDs(ICollection<int> parsedIDs)
        {
            int intValue;
            var stringIDs = ImportIDs.Split(new[] { ' ', ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            var intIDs = stringIDs.Where(x => int.TryParse(x, out intValue)).Select(int.Parse).ToList();

            parsedIDs.Clear();
            foreach (var id in intIDs)
            {
                parsedIDs.Add(id);
            }

            if (parsedIDs.Count == 0 || stringIDs.Any(x => int.TryParse(x, out intValue) == false))
            {
                MessageBox.Show(Resources.GetWorkItems_Error_WrongIDs, Resources.MessageBox_Title, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return false;
            }

            return true;
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
                OnPropertyChanged(nameof(CanImport));
            }
        }

        #endregion Private methods
        #region Private event handler methods

        /// <summary>
        /// Background method to find the work items from TFS.
        /// </summary>
        private void DoFind()
        {
            SyncServiceTrace.D(Resources.FindWorkItems);
            //var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(WordDocument);
            var config = DocumentModel.Configuration;
            var tfsService = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(DocumentModel.TfsServer, DocumentModel.TfsProject, _queryConfiguration, config);
            var wordAdapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(DocumentModel.WordDocument, DocumentModel.Configuration);
            wordAdapter.Open(null);
            tfsService.Open(null);

            EnsureTfsLazyLoadingFinished(tfsService);

            _cancellationTokenSource.Cancel();
            _cancellationTokenSource = new CancellationTokenSource();
            var token = _cancellationTokenSource.Token;

            FoundWorkItems = (from wi in tfsService.WorkItems
                              select new DataItemModel<SynchronizedWorkItemViewModel>(new SynchronizedWorkItemViewModel(wi, null, wordAdapter, DocumentModel, token))
                              {
                                  IsChecked = true
                              }).ToList();
            var foundIds = string.Join(",", FoundWorkItems.Select(w => w.Item.Id.ToString(CultureInfo.InvariantCulture)));
            SyncServiceTrace.I("Found work items: {0}", string.Equals(string.Empty, foundIds) ? "<none>" : foundIds);

            var colors = tfsService.WorkItemColors;
            foreach (var foundItem in FoundWorkItems)
            {
                foundItem.Item.CanShowColor = false;
                if (colors != null && colors.ContainsKey(foundItem.Item.TfsItem.WorkItemType))
                {
                    foundItem.Item.WorkItemColor = $"#{colors[foundItem.Item.TfsItem.WorkItemType].PrimaryColor}";
                    foundItem.Item.CanShowColor = true;
                }
            }

            var allWorkItems = tfsService.AvailableWorkItemsFromTFS;
            var idsOfFilteredWorkItems = FoundWorkItems.Select(w => w.Item.Id);
            var amountOfFilteredWorkItems = FoundWorkItems.Count();
            ShowResultsOfComparison(allWorkItems, tfsService.WorkItems.Count(), idsOfFilteredWorkItems);
        }

        /// <summary>
        /// Background method to compare total amount of work items and filtered amount of work items.
        /// </summary>
        public SortedDictionary<string, int> CompareCountsOfWorkItems(Dictionary<int, string> idsAndTypesOfAllWorkItems, int amountOfFilteredWorkItems, IEnumerable<int> idsOfFilteredWorkItems)
        {
            var listOfWorkItems = new SortedDictionary<string, int>();
            IsExistNotShownWorkItems = false;

            if (idsAndTypesOfAllWorkItems.Count != amountOfFilteredWorkItems)
            {
                foreach (var item in idsAndTypesOfAllWorkItems)
                {

                    if (!idsOfFilteredWorkItems.Contains<int>(item.Key))
                    {
                        if (!listOfWorkItems.Keys.Contains(item.Value))
                        {
                            listOfWorkItems.Add(item.Value, 1);
                        }
                        else
                        {
                            listOfWorkItems[item.Value]++;
                        }
                    }


                }
            }
            return listOfWorkItems;
        }
        /// <summary>
        /// Background method that show results of comparison.
        /// </summary>
        private void ShowResultsOfComparison(ICollection<WorkItem> allWorkItems, int amountOfFilteredWorkItems, IEnumerable<int> idsOfFilteredWorkItems)
        {
            Dictionary<int, string> idsAndTypesOfAllWorkItems = allWorkItems.ToDictionary(f => f.Id, f => f.Type.Name);
            var resultsOfComparison = CompareCountsOfWorkItems(idsAndTypesOfAllWorkItems, amountOfFilteredWorkItems, idsOfFilteredWorkItems); //call the method that contains the code above
                                                                                                                                              // if no work item has been filtered, the label should not be shown
            if (resultsOfComparison.Keys.Count <= 0)
            {
                IsExistNotShownWorkItems = false;
            }
            else
            {
                var list = new List<string>();
                foreach (var kvp in resultsOfComparison)
                {
                    list.Add(kvp.Value + " x " + kvp.Key);
                }
                var resultOfComparing = string.Join("\n", list);
                ResultOfComparing = resultOfComparing;
                var number = allWorkItems.Count() - amountOfFilteredWorkItems;
                AmountOfWorkItems = string.Format(Resources.AmountOfWorkItemsNotShown, number);
                IsExistNotShownWorkItems = true;
                SynchronizationState = SynchronizationState.Unknown;
                OnPropertyChanged(nameof(SynchronizationState));
                OnPropertyChanged(nameof(ResultOfComparing));
                OnPropertyChanged(nameof(AmountOfWorkItems));
            }
            OnPropertyChanged(nameof(IsExistNotShownWorkItems));

        }


        /// <summary>
        /// This beautiful piece of code ensures, that the TFS API has finished creating its field collection.
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

        /// <summary>
        /// Background method to import the work items from TFS.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A <see cref="DoWorkEventArgs"/> that contains the event data.</param>
        private void DoImport(object sender, DoWorkEventArgs e)
        {
            SyncServiceTrace.I(Resources.ImportWorkItems);
            var workItemSyncService = SyncServiceFactory.GetService<IWorkItemSyncService>();
            var configService = SyncServiceFactory.GetService<IConfigurationService>();

            // save query configuration to document
            DocumentModel.SaveQueryConfiguration(_queryConfiguration);
            DocumentModel.Save();

            // set actual selected configuration
            var configuration = configService.GetConfiguration(WordDocument);

            var ignoreFormatting = configuration.IgnoreFormatting;
            var conflictOverwrite = configuration.ConflictOverwrite;
            configuration.ActivateMapping(DocumentModel.MappingShowName);
            DocumentModel.Configuration.IgnoreFormatting = ignoreFormatting;
            configuration.ConflictOverwrite = conflictOverwrite;
            configuration.IgnoreFormatting = ignoreFormatting;

            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(WordDocument);
            var source = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(DocumentModel.TfsServer, DocumentModel.TfsProject, null, config);
            _importWordAdapter = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(WordDocument, configuration);

            _importWordAdapter.PrepareDocumentForLongTermOperation();
            //////////
            _importWordAdapter.ProcessOperations(configuration.PreOperations);
            var importWorkItems = (from wim in FoundWorkItems
                                   where wim.IsChecked
                                   select wim.Item).ToList();

            // search for already existing work items in word and ask whether to overide them
            if (!(source.Open(importWorkItems.Select(x => x.TfsItem.Id).ToArray()) && _importWordAdapter.Open(null)))
            {
                return;
            }

            if (importWorkItems.Select(x => x.TfsItem.Id).Intersect(_importWordAdapter.WorkItems.Select(x => x.Id)).Any())
            {
                SyncServiceTrace.W(Resources.TFSExport_ExistingWorkItems);
                if (
                    System.Windows.Forms.MessageBox.Show((IWin32Window)WordRibbon,
                                                         Resources.TFSExport_ExistingWorkItems,
                                                         Resources.MessageBox_Title,
                                                         MessageBoxButtons.YesNo,
                                                         MessageBoxIcon.Question,
                                                         MessageBoxDefaultButton.Button2, 0) == DialogResult.No)
                {
                    return;
                }
            }

            workItemSyncService.Refresh(source, _importWordAdapter, importWorkItems.Select(x => x.TfsItem), configuration);
            _importWordAdapter.ProcessOperations(configuration.PostOperations);
        }

        /// <summary>
        /// Occurs when the background operation has completed, has been canceled, or has raised an exception.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A <see cref="RunWorkerCompletedEventArgs"/> that contains the event data.</param>
        private void ImportFinished(object sender, RunWorkerCompletedEventArgs e)
        {
            e.Error.NotifyIfException("Failed to import work items.");

            IsImporting = false;
            _importWordAdapter.UndoPreparationsDocumentForLongTermOperation();
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.HideProgress();

            // enable all other functionality
            WordRibbon?.EnableAllControls(DocumentModel);
            SyncServiceTrace.I(Resources.ImportFinished);
        }

        #endregion Private event handler methods
    }
}