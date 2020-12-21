#region Usings
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Contracts.Word;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Model.WindowModelBase;
using AIT.TFS.SyncService.Contracts.ProgressService;
using Microsoft.Office.Interop.Word;
using AIT.TFS.SyncService.Contracts.Adapter;
#endregion

namespace AIT.TFS.SyncService.Model.TestReport
{
    /// <summary>
    /// The class implements base model for test report models
    /// </summary>
    public abstract class TestReportModel : ExtBaseModel
    {
        #region Private constants

        private readonly char[] _constKeyValuePartsDelimiter = new[] { '\n' };
        private readonly char[] _constValueDelimiter = new[] { '\r' };

        private const string ConstSelectedServerBuild = "SelectedServerBuild";
        private const string ConstSelectedTestBuild = "SelectedTestBuild";
        private const string ConstSelectedTestSuite = "SelectedTestSuite";
        private const string ConstSelectedTestConfiguration = "SelectedTestConfiguration";
        private const string ConstCreateDocumentStructure = "CreateDocumentStructure";
        private const string ConstIncludeTestConfigurations = "IncludeTestConfigurations";
        private const string ConstSkipLevels = "SkipLevels";
        private const string ConstSelectedDocumentStructureType = "SelectedDocumentStructureType";
        private const string ConstSelectedConfigurationPositionType = "SelectedConfigurationPositionType";

        private const string ConstSelectedTestCaseSortType = "SelectedSelectedTestCaseSortType";
        private const string ConstIncludeOnlyMostRecentTestResult = "IncludeOnlyMostRecentTestResult";
        private const string ConstIncludeOnlyMostRecentTestResultForAllConfigurations = "IncludeOnlyMostRecentTestResultForAllConfigurations";

        #endregion Private constants

        #region Private fields

        private readonly IDictionary<string, string> _storedData;
        private readonly string _modelKey;
        private readonly List<BackgroundWorker> _backgroundWorkers = new List<BackgroundWorker>();
        private bool _generatingActive;
        private bool _operationActive;
        /// <summary>
        /// It's possible that on initialization time more background worker runs. Here is the first state of operation active.
        /// </summary>
        private bool? _oldOperationActive;

        protected ITestReportingProgressCancellationService TestReportingProgressCancellationService;

        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TestResultReportModel"/> class.
        /// </summary>
        /// <param name="syncServiceDocumentModel">The <see cref="ISyncServiceDocumentModel"/> to obtain document settings.</param>
        /// <param name="viewDispatcher"> Instance of <see cref="IViewDispatcher"/> of associated view.</param>
        /// <param name="testAdapter"><see cref="ITfsTestAdapter"/> to examine test information.</param>
        /// <param name="wordRibbon">Interface to the word ribbon</param>
        /// <param name="testReportingProgressCancellationService">The <see cref="ITestReportingProgressCancellationService"/> indicating that a cancellation is requested.</param>
        protected TestReportModel(ISyncServiceDocumentModel syncServiceDocumentModel, IViewDispatcher viewDispatcher, ITfsTestAdapter testAdapter, IWordRibbon wordRibbon, ITestReportingProgressCancellationService testReportingProgressCancellationService)
        {
            if (syncServiceDocumentModel == null)
                throw new ArgumentNullException("syncServiceDocumentModel");
            if (testAdapter == null)
                throw new ArgumentNullException("testAdapter");
            if (wordRibbon == null)
                throw new ArgumentNullException("wordRibbon");
            if (testReportingProgressCancellationService == null)
                throw new ArgumentNullException("testReportingProgressCancellationService");
            if (viewDispatcher == null)
                throw new ArgumentNullException("viewDispatcher");

            SyncServiceDocumentModel = syncServiceDocumentModel;
            WordRibbon = wordRibbon;
            TestAdapter = testAdapter;
            ViewDispatcher = viewDispatcher;

            Document = syncServiceDocumentModel.WordDocument as Document;

            _modelKey = GetType().Name;
            _storedData = new Dictionary<string, string>();
            ParseStoredData(SyncServiceDocumentModel.ReadTestReportData(_modelKey));

            _generatingActive = SyncServiceDocumentModel.TestReportRunning;
            _operationActive = SyncServiceDocumentModel.TestReportRunning;

            TestReportingProgressCancellationService = testReportingProgressCancellationService;
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="TestResultReportModel"/> class for the console extension
        /// </summary>
        /// <param name="syncServiceDocumentModel">The <see cref="ISyncServiceDocumentModel"/> to obtain document settings.</param>
        /// <param name="testAdapter"><see cref="ITfsTestAdapter"/> to examine test information.</param>
        /// <param name="testReportingProgressCancellationService">The <see cref="ITestReportingProgressCancellationService"/> indicating that a cancellation is requested.</param>
        protected TestReportModel(ISyncServiceDocumentModel syncServiceDocumentModel, ITfsTestAdapter testAdapter, ITestReportingProgressCancellationService testReportingProgressCancellationService)
        {
            if (syncServiceDocumentModel == null)
                throw new ArgumentNullException("syncServiceDocumentModel");
            if (testAdapter == null)
                throw new ArgumentNullException("testAdapter");
            if (testReportingProgressCancellationService == null)
                throw new ArgumentNullException("testReportingProgressCancellationService");

            SyncServiceDocumentModel = syncServiceDocumentModel;
            TestAdapter = testAdapter;

            Document = syncServiceDocumentModel.WordDocument as Document;

            _modelKey = GetType().Name;
            _storedData = new Dictionary<string, string>();
            ParseStoredData(SyncServiceDocumentModel.ReadTestReportData(_modelKey));

            _generatingActive = SyncServiceDocumentModel.TestReportRunning;
            _operationActive = SyncServiceDocumentModel.TestReportRunning;
            TestReportingProgressCancellationService = testReportingProgressCancellationService;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets interface to the word ribbon.
        /// </summary>
        public IWordRibbon WordRibbon { get; private set; }

        /// <summary>
        /// Gets the <see cref="ISyncServiceDocumentModel"/>
        /// </summary>
        public ISyncServiceDocumentModel SyncServiceDocumentModel { get; private set; }

        /// <summary>
        /// Gets <see cref="ViewDispatcher"/> of associated view.
        /// </summary>
        public IViewDispatcher ViewDispatcher { get; private set; }

        /// <summary>
        /// Gets or sets the <see cref="ITfsTestAdapter"/> to obtain test information.
        /// </summary>
        public ITfsTestAdapter TestAdapter { get; set; }

        /// <summary>
        /// Gets or sets the associated <see cref="Document"/>.
        /// </summary>
        public Document Document { get; set; }

        /// <summary>
        /// Gets the <see cref="ViewCommand"/> for 'Create Report'.
        /// </summary>
        public ViewCommand CreateReportCommand { get; protected set; }

        #endregion Public properties

        #region Protected properties - stored data related

        /// <summary>
        /// Gets or sets the persisted 'selected server build'.
        /// </summary>
        protected string StoredSelectedServerBuild
        {
            get { return GetStoredProperty(ConstSelectedServerBuild); }
            set { _storedData[ConstSelectedServerBuild] = value; }
        }

        /// <summary>
        /// Gets or sets the persisted 'selected test plan'.
        /// </summary>
        protected string StoredSelectedTestPlan
        {
            get { return GetStoredProperty(ConstSelectedTestBuild); }
            set { _storedData[ConstSelectedTestBuild] = value; }
        }

        /// <summary>
        /// Gets or sets the persisted 'selected test suite'.
        /// </summary>
        protected string StoredSelectedTestSuite
        {
            get { return GetStoredProperty(ConstSelectedTestSuite); }
            set { _storedData[ConstSelectedTestSuite] = value; }
        }

        /// <summary>
        /// Gets or sets the persisted 'selected test configuration'.
        /// </summary>
        protected string StoredSelectedTestConfiguration
        {
            get { return GetStoredProperty(ConstSelectedTestConfiguration); }
            set { _storedData[ConstSelectedTestConfiguration] = value; }
        }

        /// <summary>
        /// Gets or sets the persisted 'create document structure'.
        /// </summary>
        protected bool StoredCreateDocumentStructure
        {
            get
            {
                return string.IsNullOrEmpty(GetStoredProperty(ConstCreateDocumentStructure));
            }
            set
            {
                _storedData[ConstCreateDocumentStructure] = value ? string.Empty : "false";
            }
        }

        /// <summary>
        /// Gets or sets the persisted 'include test configurations'.
        /// </summary>
        protected bool StoredIncludeTestConfigurations
        {
            get
            {
                return string.IsNullOrEmpty(GetStoredProperty(ConstIncludeTestConfigurations));
            }
            set
            {
                _storedData[ConstIncludeTestConfigurations] = value ? string.Empty : "false";
            }
        }

        /// <summary>
        /// Gets or sets the persisted 'skip levels'.
        /// </summary>
        protected int StoredSkipLevels
        {
            get
            {
                var value = GetStoredProperty(ConstSkipLevels);
                int intValue;
                if (!int.TryParse(value, out intValue))
                    intValue = 2;
                return intValue;
            }
            set
            {
                _storedData[ConstSkipLevels] = value.ToString(CultureInfo.CurrentCulture);
            }
        }

        /// <summary>
        /// Gets or sets the persisted 'selected document structure type'.
        /// </summary>
        protected DocumentStructureType StoredSelectedDocumentStructureType
        {
            get
            {
                var value = GetStoredProperty(ConstSelectedDocumentStructureType);
                DocumentStructureType type;
                if (!Enum.TryParse(value, out type))
                    type = DocumentStructureType.IterationPath;
                return type;
            }
            set
            {
                _storedData[ConstSelectedDocumentStructureType] = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the persisted 'configuration position type'.
        /// </summary>
        protected ConfigurationPositionType StoredSelectedConfigurationPositionType
        {
            get
            {
                var value = GetStoredProperty(ConstSelectedConfigurationPositionType);
                ConfigurationPositionType type;
                if (!Enum.TryParse(value, out type))
                    type = ConfigurationPositionType.AboveTestPlan;
                return type;
            }
            set
            {
                _storedData[ConstSelectedConfigurationPositionType] = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the persisted 'selected test case sort type'.
        /// </summary>
        protected TestCaseSortType StoredSelectedTestCaseSortType
        {
            get
            {
                var value = GetStoredProperty(ConstSelectedTestCaseSortType);
                TestCaseSortType type;
                if (!Enum.TryParse(value, out type))
                    type = TestCaseSortType.None;
                return type;
            }
            set
            {
                _storedData[ConstSelectedTestCaseSortType] = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the persisted value that determines if only the most recent test results are shown
        /// </summary>
        protected bool StoredIncludeOnlyMostRecentTestResult
        {
            get
            {
                var value = GetStoredProperty(ConstIncludeOnlyMostRecentTestResult);
                bool result;
                if (!bool.TryParse(value, out result))
                {
                    return false;
                }
                return result;
            }
            set
            {
                _storedData[ConstIncludeOnlyMostRecentTestResult] = value.ToString();
            }
        }


        /// <summary>
        /// Gets or sets the persisted value that determines if only the most recent test results are shown
        /// </summary>
        protected bool StoredIncludeOnlyMostRecentTestResultForAllConfigurations
        {
            get
            {
                var value = GetStoredProperty(ConstIncludeOnlyMostRecentTestResultForAllConfigurations);
                bool result;
                if (!bool.TryParse(value, out result))
                {
                    return false;
                }
                return result;
            }
            set
            {
                _storedData[ConstIncludeOnlyMostRecentTestResultForAllConfigurations] = value.ToString();
            }
        }

        #endregion Protected properties - stored data related

        #region Public binding properties

        /// <summary>
        /// Gets or sets the flag telling if the generating (create document) is active.
        /// </summary>
        public bool GeneratingActive
        {
            get
            {
                return _generatingActive;
            }
            set
            {
                if (_generatingActive == value)
                {
                    return;
                }
                _generatingActive = value;
                if (_generatingActive)
                {
                    // Reset old showed messages.
                    WordRibbon.ResetBeforeOperation(SyncServiceDocumentModel);
                    // Disable all other buttons during the operation
                    WordRibbon.DisableAllControls(SyncServiceDocumentModel);
                }
                else
                {
                    WordRibbon.EnableAllControls(SyncServiceDocumentModel);
                }
                OnPropertyChanged(nameof(GeneratingActive));
            }
        }

        /// <summary>
        /// Gets or sets the flag telling if the operation (create document) is active.
        /// </summary>
        public bool OperationActive
        {
            get
            {
                return _operationActive;
            }
            set
            {
                if (_operationActive == value)
                {
                    return;
                }
                _operationActive = value;
                OnPropertyChanged(nameof(OperationActive));
            }
        }

        #endregion Public binding properties

        #region Public methods

        /// <summary>
        /// The method is called after the visibility of panel is changed.
        /// </summary>
        /// <param name="show"><c>true</c> - associated view is shown.</param>
        public void VisibilityChanged(bool show)
        {
            if (show)
            {
                SyncServiceDocumentModel.TestReportDataChanged += HandleTestReportDataChangedEvent;
                return;
            }
            else
            {
                SyncServiceDocumentModel.TestReportDataChanged -= HandleTestReportDataChangedEvent;
            }

            // Hide: cancel all operations
            lock (_backgroundWorkers)
            {
                if (_backgroundWorkers.Count == 0)
                {
                    return;
                }
                foreach (var bw in _backgroundWorkers)
                {
                    bw.CancelAsync();
                }
            }
        }

        #endregion Public methods

        #region Protected methods

        /// <summary>
        /// The method starts background worker with given action. Long operation must check cancelation pending.
        /// </summary>
        /// <param name="workerForGenerating">Action is for test report generating.</param>
        /// <param name="workerAction">Action to run in background worker.</param>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Disposed in 'RunWorkerCompleted'")]
        protected void StartBackgroundWorker(bool workerForGenerating, Action workerAction)
        {
            lock (_backgroundWorkers)
            {
                // Store the old value of operation active only if no background worker is running
                if (_backgroundWorkers.Count == 0)
                {
                    _oldOperationActive = OperationActive;
                }
                var backgroundWorker = new BackgroundWorker { WorkerSupportsCancellation = true };
                backgroundWorker.DoWork += (sender, e) =>
                {
                    if (workerForGenerating)
                    {
                        GeneratingActive = true;
                    }
                    OperationActive = true;
                    workerAction();
                };
                backgroundWorker.RunWorkerCompleted += (sender, e) =>
                {
                    e.Error.NotifyIfException("An error occurred.");
                    lock (_backgroundWorkers)
                    {
                        if (workerForGenerating)
                        {
                            GeneratingActive = false;
                        }
                        // Set the old value of operation active only if last bacground worker is stopping
                        if (_backgroundWorkers.Count == 1 && _oldOperationActive.HasValue)
                        {
                            OperationActive = _oldOperationActive.Value;
                            _oldOperationActive = null;
                        }
                        backgroundWorker.Dispose();
                        Debug.Assert(_backgroundWorkers.Contains(backgroundWorker), "Somethin is wrong. Background worker is noct in list.");
                        _backgroundWorkers.Remove(backgroundWorker);
                    }
                };
                _backgroundWorkers.Add(backgroundWorker);
                backgroundWorker.RunWorkerAsync();
            }
        }

        /// <summary>
        /// The method determines if the cancelation of background worker is pending.
        /// </summary>
        /// <returns><c>true</c> - cancelation is pending. Otherwise <c>false</c>.</returns>
        protected bool CancellationPending()
        {
            lock (_backgroundWorkers)
            {
                if (_backgroundWorkers.Count == 0)
                {
                    return false;
                }
                return _backgroundWorkers[0].CancellationPending;
            }
        }

        /// <summary>
        /// The method stores data to made this data persisted. Persisted data are read in constructor.
        /// </summary>
        protected void StoreReportData()
        {
            SyncServiceDocumentModel.WriteTestReportData(_modelKey, PrepareDataToStore());
        }

        #endregion Protected methods

        #region Private methods - stored data related

        /// <summary>
        /// Gets the value for given property key.
        /// </summary>
        /// <param name="propertyKey">Property key to get the value for.</param>
        /// <returns>Stored value or <c>null</c> if not stored.</returns>
        private string GetStoredProperty(string propertyKey)
        {
            if (_storedData.ContainsKey(propertyKey))
            {
                return _storedData[propertyKey];
            }
            return null;
        }

        /// <summary>
        /// The method parses the stored data. See <see cref="PrepareDataToStore"/>.
        /// </summary>
        /// <param name="storedData">Stored data to parse.</param>
        private void ParseStoredData(string storedData)
        {
            if (string.IsNullOrEmpty(storedData))
            {
                return;
            }
            var keyValueParts = storedData.Split(_constKeyValuePartsDelimiter);
            foreach (var keyValuePart in keyValueParts)
            {
                var keyValue = keyValuePart.Split(_constValueDelimiter);
                if (keyValue.Length == 2)
                {
                    _storedData[keyValue[0]] = keyValue[1];
                }
            }
        }

        /// <summary>
        /// The method creates data to store. See <see cref="ParseStoredData"/>.
        /// </summary>
        /// <returns>Data ready to store.</returns>
        private string PrepareDataToStore()
        {
            var dataList = new List<string>();
            foreach (var keyValuePair in _storedData)
            {
                dataList.Add(string.Join(new string(_constValueDelimiter), keyValuePair.Key, keyValuePair.Value));
            }
            return string.Join(new string(_constKeyValuePartsDelimiter), dataList.ToArray()); ;
        }

        #endregion Private methods - stored data related

        #region Private event handler methods

        /// <summary>
        /// Occurs if any test result data in <see cref="ISyncServiceModel"/> was changed.
        /// </summary>
        /// <param name="sender">Sender of the event.</param>
        /// <param name="e">Empty event data.</param>
        private void HandleTestReportDataChangedEvent(object sender, SyncServiceEventArgs e)
        {
            if (e.DocumentName != Document.Name || e.DocumentFullName != Document.FullName)
            {
                return;
            }

            ViewDispatcher.Invoke(() => CreateReportCommand.CallEventCanExecuteChanged());
            OperationActive = SyncServiceDocumentModel.TestReportRunning;
        }

        #endregion Private event handler methods
    }
}
