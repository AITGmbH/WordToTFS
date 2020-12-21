#region Usings
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.WindowModelBase;
using AIT.TFS.SyncService.Service.Utils;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
#endregion

namespace AIT.TFS.SyncService.Model
{
    /// <summary>
    /// Class that represents the synchronization state between the Word and the TFS version of a work item.
    /// </summary>
    public class SynchronizedWorkItemViewModel : ExtBaseModel
    {
        #region Private fields
        private static readonly object _dontWaitAfterHtmlAccessLock = new object();
        private readonly CancellationToken _cancellationToken;
        private readonly ISyncServiceDocumentModel _documentModel;
        private readonly IWorkItemSyncAdapter _wordAdapter;
        private SynchronizationState _synchronizationState;
        private IEnumerable<string> _fields;
        #endregion
        #region Public properties
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
        /// TFS version of this work item.
        /// </summary>
        public IWorkItem TfsItem
        {
            get;
            private set;
        }

        /// <summary>
        /// TFS version of this work item.
        /// </summary>
        public IWorkItem WordItem
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets a list of fields that have changes
        /// </summary>
        public IEnumerable<string> Fields
        {
            get
            {
                return _fields;
            }
            private set
            {
                _fields = value;
                TriggerPropertyChanged(nameof(Fields));
            }
        }

        /// <summary>
        /// Gets the title of the work item in Word. If the work item is not imported, gets the title of the work item on the server.
        /// </summary>
        public string Title
        {
            get
            {
                return TfsItem != null ? TfsItem.Title : WordItem.Title;
            }
        }

        /// <summary>
        /// Gets the id of this work item.
        /// </summary>
        public int Id
        {
            get
            {
                return TfsItem == null ? WordItem.Id : TfsItem.Id;
            }
        }

        /// <summary>
        /// Color of work item.
        /// </summary>
        public string WorkItemColor { get; set; }

        /// <summary>
        /// Indicate if color can be shown.
        /// </summary>
        public bool CanShowColor { get; set; }

        #endregion
        #region Public methods
        /// <summary>
        /// Creates a new instance of this class.
        /// </summary>
        /// <param name="tfsWorkItem">The TFS version of this work item.</param>
        /// <param name="wordWorkItem">The Word version of this work item.</param>
        /// <param name="wordAdapter">The Word adapter from which to load.</param>
        /// <param name="documentModel">The model of the current document for which to monitor sync state.</param>
        /// <param name="cancellationToken">Token used to cancel synchronization state query.</param>
        public SynchronizedWorkItemViewModel(IWorkItem tfsWorkItem, IWorkItem wordWorkItem, IWorkItemSyncAdapter wordAdapter, ISyncServiceDocumentModel documentModel, CancellationToken? cancellationToken)
        {
            if (cancellationToken.HasValue)
            {
                _cancellationToken = cancellationToken.Value;
            }

            _wordAdapter = wordAdapter;
            _documentModel = documentModel;

            TfsItem = tfsWorkItem;
            WordItem = wordWorkItem;

            SynchronizationState = SynchronizationState.Unknown;

            Task.Factory.StartNew(Load, _cancellationToken).ContinueWith(CheckIfCancelled);
        }

        /// <summary>
        /// Refreshes work items and synchronization state
        /// </summary>
        public void Refresh()
        {
            if (TfsItem == null)
            {
                LoadTfsItem();
            }
            else
            {
                TfsItem.Refresh();
            }

            if (WordItem == null)
            {
                LoadWordItem();
            }
            else
            {
                WordItem.Refresh();
            }

            RefreshState();
        }

        /// <summary>
        /// Load items and refresh state
        /// </summary>
        private void Load()
        {
            // Give Word time to refresh its UI with "Unknown Status" entries
            Thread.Sleep(500);
            if (WordItem == null)
            {
                LoadWordItem();
            }
            else
            {
                LoadTfsItem();
            }

            RefreshState();
        }

        /// <summary>
        /// Set state to aborted if cancelled
        /// </summary>
        private void CheckIfCancelled(Task task)
        {
            if (task.IsCanceled)
            {
                SynchronizationState = SynchronizationState.Aborted;
            }
        }

        /// <summary>
        /// Load work item from TFS
        /// </summary>
        private void LoadTfsItem()
        {
            if (Id == 0)
            {
                return;
            }

            var source = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(_documentModel.TfsServer, _documentModel.TfsProject, null, _documentModel.Configuration);

            if (!source.Open(new[] { Id }))
            {
                return;
            }

            TfsItem = source.WorkItems.Find(Id);
        }

        /// <summary>
        /// Load work item from Word adapter
        /// </summary>
        private void LoadWordItem()
        {
            if (_wordAdapter.Open(null) == false) return;
            WordItem = _wordAdapter.WorkItems.Find(Id);
        }

        /// <summary>
        /// Assign state based on field value comparisons
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "Its fine and understandable")]
        private void RefreshState()
        {
            if (WordItem != null && WordItem.IsNew)
            {
                Fields = WordItem.Fields.Select(x => x.ReferenceName);
                SynchronizationState = SynchronizationState.New;
                return;
            }

            if (WordItem == null)
            {
                Fields = TfsItem.Fields.Select(x => x.ReferenceName);
                SynchronizationState = SynchronizationState.NotImported;
                return;
            }

            if (WordItem.Revision < 1 || WordItem.Revision > TfsItem.Revision)
            {
                // broken revision field in word
                return;
            }

            var changedTfsFields = SynchronizationInformationHelper.GetChangedTfsFields(WordItem, TfsItem, _documentModel.Configuration.IgnoreFormatting);
            ThrowIfCancelled();


            var refreshableTfsFields = SynchronizationInformationHelper.GetRefreshableTfsFields(changedTfsFields, TfsItem);
            Collection<IField> changedWordFields;

            // there is another lock that prevents multiple threads from accessing the clipboard at the same time.
            // I don't want all sync tasks to wait at that lock because then I cannot stop them without reading from word.
            lock (_dontWaitAfterHtmlAccessLock)
            {
                ThrowIfCancelled();
                changedWordFields = SynchronizationInformationHelper.GetChangedWordFields(WordItem, TfsItem, _documentModel.Configuration.IgnoreFormatting);
            }

            var publishableWordFields = SynchronizationInformationHelper.GetPublishableWordFields(changedWordFields, TfsItem);
            var divergedFields = SynchronizationInformationHelper.GetDivergedFields(publishableWordFields, changedTfsFields);

            if (!publishableWordFields.Any() && refreshableTfsFields.Any())
            {
                Fields = refreshableTfsFields.Select(x => x.ReferenceName);
                SynchronizationState = SynchronizationState.Outdated;
            }

            if (!publishableWordFields.Any() && !refreshableTfsFields.Any())
            {
                // Hidden fields can never be updated - no point nagging about differing values
                Fields = changedWordFields.Select(x => x.ReferenceName);
                SynchronizationState = changedWordFields.Any(x => x.Configuration.IsMapped) ? SynchronizationState.Differing : SynchronizationState.UpToDate;
            }

            if (publishableWordFields.Any() && !refreshableTfsFields.Any() && !divergedFields.Any())
            {
                Fields = publishableWordFields.Select(x => x.ReferenceName);
                SynchronizationState = SynchronizationState.Dirty;
            }

            if (publishableWordFields.Any() && divergedFields.Any())
            {
                Fields = refreshableTfsFields.Select(x => x.ReferenceName).Concat(publishableWordFields.Select(x => x.ReferenceName)).Distinct();
                SynchronizationState = SynchronizationState.DivergedWithConflicts;
            }

            if (publishableWordFields.Any() && changedTfsFields.Any() && !divergedFields.Any())
            {
                Fields = refreshableTfsFields.Select(x => x.ReferenceName).Concat(publishableWordFields.Select(x => x.ReferenceName)).Distinct();
                SynchronizationState = SynchronizationState.DivergedWithoutConflicts;
            }
        }

        /// <summary>
        /// Set state to aborted and stop task
        /// </summary>
        private void ThrowIfCancelled()
        {
            if (_cancellationToken.IsCancellationRequested)
            {
                _cancellationToken.ThrowIfCancellationRequested();
            }
        }
        #endregion
    }
}
