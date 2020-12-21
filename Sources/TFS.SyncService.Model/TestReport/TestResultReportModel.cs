#region Usings
using System;
using System.Collections.Generic;
using System.Linq;
using AIT.TFS.SyncService.Contracts.BuildCenter;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Contracts.Word;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Model.Properties;
using AIT.TFS.SyncService.Service.InfoStorage;
using Microsoft.Office.Interop.Word;
using System.Collections.ObjectModel;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Contracts.Adapter;
#endregion

namespace AIT.TFS.SyncService.Model.TestReport
{
    /// <summary>
    /// Model class for TestResultReport.
    /// </summary>
    public sealed class TestResultReportModel : TestReportModel
    {
        #region Private fields

        private readonly ObservableCollection<ITfsTestSuite> _treeViewTestSuites = new ObservableCollection<ITfsTestSuite>();
        private object _selectedTreeViewItem;
        private ITfsTestSuite _selectedTestSuite;
        private IList<ITfsServerBuild> _serverBuilds;
        private ITfsServerBuild _selectedServerBuild;
        private ITfsServerBuild _lastUsedSelectedServerBuildForTestPlans = AllServerBuilds.Empty;
        private ITfsServerBuild _lastUsedSelectedServerBuildForTestConfigurations = AllServerBuilds.Empty;
        private IList<ITfsTestPlan> _testPlans;
        private ITfsTestPlan _selectedTestPlan;
        private IList<ITfsTestConfiguration> _testConfigurations;
        private ITfsTestPlan _lastUsedTestPlanForTestConfigurations;
        private ITfsTestConfiguration _selectedTestConfiguration;
        private bool _createDocumentStructure;
        private DocumentStructureType _selectedDocumentStructureType;
        private ConfigurationPositionType _selectedConfigurationPositionType;
        private int _skipLevels;
        private bool _includeTestConfigurations;
        private TestCaseSortType _selectedTestCaseSortType;
        private bool _includeOnlyMostRecentTestResult;
        private bool _includeOnlyMostRecentTestResultForAllConfigurations;
        private bool _testConfigurationInformationPrinted;
        private string _defaultTestPlanName = string.Empty;
        private string _defaultConfigurationName = string.Empty;
        private string _defaultServerBuildName = string.Empty;

        public static Dictionary<int, string> Paths = new Dictionary<int, string>();
        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TestResultReportModel"/> class.
        /// </summary>
        /// <param name="testReportingProgrgessCancellationService">Service to cancel further steps in test reporting generation.</param>
        /// <param name="syncServiceDocumentModel">The <see cref="ISyncServiceModel"/> to obtain document settings.</param>
        /// <param name="dispatcher">Dispatcher of associated view.</param>
        /// <param name="testAdapter"><see cref="ITfsTestAdapter"/> to examine test information.</param>
        /// <param name="wordRibbon">Interface to the word ribbon</param>
        public TestResultReportModel(ISyncServiceDocumentModel syncServiceDocumentModel, IViewDispatcher dispatcher,
            ITfsTestAdapter testAdapter, IWordRibbon wordRibbon, ITestReportingProgressCancellationService testReportingProgrgessCancellationService)
            : base(syncServiceDocumentModel, dispatcher, testAdapter, wordRibbon, testReportingProgrgessCancellationService)
        {
            CreateDocumentStructure = StoredCreateDocumentStructure;
            IncludeTestConfigurations = StoredIncludeTestConfigurations;

            SkipLevels = StoredSkipLevels;
            SelectedDocumentStructureType = StoredSelectedDocumentStructureType;
            SelectedTestCaseSortType = StoredSelectedTestCaseSortType;

            SelectedConfigurationPositionType = StoredSelectedConfigurationPositionType;

            IncludeOnlyMostRecentTestResult = StoredIncludeOnlyMostRecentTestResult;
            IncludeOnlyMostRecentTestResultForAllConfigurations = StoredIncludeOnlyMostRecentTestResultForAllConfigurations;

            CreateReportCommand = new ViewCommand(ExecuteCreateReportCommand, CanExecuteCreateReportCommand);

            SetTestReportDefaults(syncServiceDocumentModel.Configuration.ConfigurationTest.ConfigurationTestResult.DefaultValues);

            WordDocument = syncServiceDocumentModel.WordDocument as Document;
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="TestSpecificationReportModel"/> class for the console extension
        /// </summary>
        /// <param name="syncServiceDocumentModel">The <see cref="ISyncServiceModel"/> to obtain document settings.</param>
        /// <param name="testAdapter"><see cref="ITfsTestAdapter"/> to examine test information.</param>
        /// <param name="plan">The test plan.</param>
        /// <param name="suite"></param>
        /// <param name="documentStructure"></param>
        /// <param name="skipLevels"></param>
        /// <param name="structureType"></param>
        /// <param name="sort"></param>
        /// <param name="includeTestConfigurations"></param>
        /// <param name="position"></param>
        /// <param name="includeOnlyMostRecent"></param>
        /// <param name="recentForAllConfigurations"></param>
        /// <param name="buildName"></param>
        /// <param name="configurationName"></param>
        /// <param name="testReportingProgrgessCancellationService">Service to cancel further steps in test reporting generation.</param>
        public TestResultReportModel(ISyncServiceDocumentModel syncServiceDocumentModel, ITfsTestAdapter testAdapter, ITfsTestPlan plan, ITfsTestSuite suite, bool documentStructure, int skipLevels, DocumentStructureType structureType, TestCaseSortType sort, bool includeTestConfigurations, ConfigurationPositionType position, bool includeOnlyMostRecent, bool recentForAllConfigurations, string buildName, string configurationName, ITestReportingProgressCancellationService testReportingProgrgessCancellationService)
            : base(syncServiceDocumentModel, testAdapter, testReportingProgrgessCancellationService)
        {
            CreateDocumentStructure = documentStructure;
            SelectedServerBuild = GetServerBuilds().Where(build => build.BuildNumber == buildName).FirstOrDefault();
            SelectedTestPlan = plan;
            SelectedTestSuite = suite;
            SelectedTestConfiguration = TestConfigurations.Where(configuration => configuration.Name == configurationName).FirstOrDefault();
            IncludeTestConfigurations = includeTestConfigurations;
            SelectedConfigurationPositionType = position;

            SkipLevels = skipLevels;
            SelectedDocumentStructureType = structureType;
            SelectedTestCaseSortType = sort;

            IncludeOnlyMostRecentTestResult = includeOnlyMostRecent;
            IncludeOnlyMostRecentTestResultForAllConfigurations = recentForAllConfigurations;



            WordDocument = syncServiceDocumentModel.WordDocument as Document;
        }

        #endregion Constructors

        #region Public binding properties
        /// <summary>
        /// Gets the associated word document.
        /// </summary>
        public Document WordDocument { get; private set; }


        /// <summary>
        /// Gets the list of all available server builds.
        /// </summary>
        public IList<ITfsServerBuild> ServerBuilds
        {
            get
            {
                if (_serverBuilds == null)
                {
                    StartBackgroundWorker(false, () =>
                    {
                        _serverBuilds = new List<ITfsServerBuild> { new AllServerBuilds() };
                        try
                        {
                            foreach (var element in TestAdapter.GetAvailableServerBuilds())
                                _serverBuilds.Add(element);
                        }
                        finally
                        {
                            OnPropertyChanged(nameof(ServerBuilds));
                            if (ViewDispatcher != null)
                            {
                                ViewDispatcher.BeginInvoke(new Action(() =>
                                {
                                    // Select one server build
                                    var serverBuildToSelect = _serverBuilds[0];
                                    if (!string.IsNullOrEmpty(StoredSelectedServerBuild) && _serverBuilds.Any(build => build.BuildNumber == StoredSelectedServerBuild))
                                    {
                                        serverBuildToSelect = _serverBuilds.First(build => build.BuildNumber == StoredSelectedServerBuild);
                                    }
                                    SelectedServerBuild = serverBuildToSelect;
                                }));
                            }
                        }
                    });
                }
                return _serverBuilds;
            }
        }

        /// <summary>
        /// Gets or sets selected server build.
        /// </summary>
        public ITfsServerBuild SelectedServerBuild
        {
            get
            {
                return _selectedServerBuild;
            }
            set
            {
                if (_selectedServerBuild == value)
                    return;

                //on first value assignment use default value
                if (_selectedServerBuild == null && ViewDispatcher != null && ViewDispatcher.IsDispatching)
                {
                    var defaultBuild = ServerBuilds.FirstOrDefault(x => x.BuildNumber == _defaultServerBuildName);
                    value = defaultBuild ?? value;
                }

                _selectedServerBuild = value;
                if (_selectedServerBuild != null)
                {
                    StoredSelectedServerBuild = _selectedServerBuild.BuildNumber;
                }
                OnPropertyChanged(nameof(SelectedServerBuild));
                OnPropertyChanged(nameof(TestPlans));
                if (CreateReportCommand != null && ViewDispatcher != null)
                {
                    ViewDispatcher.Invoke(() => CreateReportCommand.CallEventCanExecuteChanged());
                }
            }
        }

        /// <summary>
        /// Gets the list of all available test plans
        /// </summary>
        public IList<ITfsTestPlan> TestPlans
        {
            get
            {
                var serverBuild = SelectedServerBuild;
                if (serverBuild != null && serverBuild.BuildNumber == AllServerBuilds.AllServerBuildsId)
                // null means use all server builds
                {
                    serverBuild = null;
                }
                if (_lastUsedSelectedServerBuildForTestPlans != serverBuild)
                {
                    _testPlans = null;
                    _lastUsedSelectedServerBuildForTestPlans = serverBuild;
                    StartBackgroundWorker(false, () =>
                    {
                        _testPlans = new List<ITfsTestPlan>(TestAdapter.GetTestPlans(serverBuild)).OrderBy(o => o.Name).ToList();
                        OnPropertyChanged(nameof(TestPlans));
                        if (_testPlans != null && _testPlans.Count > 0 && ViewDispatcher != null)
                        {
                            ViewDispatcher.Invoke(SelectTestPlans);
                        }
                        else
                        {
                            SelectedTestPlan = null;
                        }
                    });
                }
                return _testPlans;
            }
        }

        private void SelectTestPlans()
        {
            // Select one test plan
            var testPlanToSelect = _testPlans[0];
            if (!string.IsNullOrEmpty(StoredSelectedTestPlan) && _testPlans.Any(plan => plan.Name == StoredSelectedTestPlan))
            {
                testPlanToSelect = _testPlans.First(plan => plan.Name == StoredSelectedTestPlan);
            }
            SelectedTestPlan = testPlanToSelect;
        }

        /// <summary>
        /// Gets or sets selected test plan.
        /// </summary>
        public ITfsTestPlan SelectedTestPlan
        {
            get
            {
                return _selectedTestPlan;
            }
            set
            {
                if (_selectedTestPlan == value)
                {
                    return;
                }

                // on first value assignment use default value
                if (_selectedTestPlan == null && ViewDispatcher != null && ViewDispatcher.IsDispatching)
                {
                    var defaultTestPlan = TestPlans.FirstOrDefault(x => x.Name == _defaultTestPlanName);
                    value = defaultTestPlan ?? value;
                }

                _selectedTestPlan = value;
                if (_selectedTestPlan != null)
                {
                    StoredSelectedTestPlan = _selectedTestPlan.Name;
                }
                OnPropertyChanged(nameof(SelectedTestPlan));
                OnPropertyChanged(nameof(TestConfigurations));
                if (ViewDispatcher != null)
                {
                    if (CreateReportCommand != null)
                    {
                        ViewDispatcher.Invoke(() => CreateReportCommand.CallEventCanExecuteChanged());
                    }

                    ViewDispatcher.Invoke(() => _treeViewTestSuites.Clear());
                }
                // Add the selected test plan to the collection for tree view if any selected
                if (_selectedTestPlan != null)
                {
                    _treeViewTestSuites.Add(_selectedTestPlan.RootTestSuite);
                }
                // No item in tree view selected
                SelectedTreeViewItem = null;
                if (CreateReportCommand != null && ViewDispatcher != null)
                {
                    ViewDispatcher.Invoke(() => CreateReportCommand.CallEventCanExecuteChanged());
                    // Call begin invoke to select default test suite in tree view
                    ViewDispatcher.BeginInvoke(
                        new Action(() =>
                        {
                            if (_selectedTestPlan != null)
                            {
                                var testSuiteSearcher = new TestSuiteSearcher(_selectedTestPlan);
                                ITfsTestSuite testSuiteToSelect;

                                try
                                {
                                    testSuiteToSelect = StoredSelectedTestSuite == null ? _selectedTestPlan.RootTestSuite : testSuiteSearcher.SearchTestSuiteWithinTestPlan(StoredSelectedTestSuite);
                                }
                                catch (ArgumentException e)
                                {
                                    testSuiteToSelect = _selectedTestPlan.RootTestSuite;
                                    SyncServiceTrace.I(string.Format(Resources.ExceptionMessage, e.Message, e.StackTrace));
                                    SyncServiceTrace.I(string.Format(Resources.SuiteNotFound, StoredSelectedTestSuite, SelectedTestPlan.Name));
                                }

                                SelectedTreeViewItem = testSuiteToSelect;
                                CreateReportCommand.CallEventCanExecuteChanged();
                            }
                        })
                    );
                }
            }
        }

        /// <summary>
        /// Gets the list of all test plans to show in tree.
        /// </summary>
        public ObservableCollection<ITfsTestSuite> TreeViewTestSuites
        {
            get { return _treeViewTestSuites; }
        }

        /// <summary>
        /// Gets or sets selected item in tree view - either <see cref="ITfsTestPlan"/> or <see cref="ITfsTestSuite"/>.
        /// </summary>
        public object SelectedTreeViewItem
        {
            get { return _selectedTreeViewItem; }
            set
            {
                if (_selectedTreeViewItem == value)
                {
                    return;
                }
                _selectedTreeViewItem = value;
                SelectedTestSuite = _selectedTreeViewItem as ITfsTestSuite;
                OnPropertyChanged(nameof(SelectedTreeViewItem));
            }
        }

        /// <summary>
        /// Gets or sets selected test suite - <see cref="ITfsTestSuite"/>.
        /// </summary>
        public ITfsTestSuite SelectedTestSuite
        {
            get { return _selectedTestSuite; }
            set
            {
                if (_selectedTestSuite == value)
                {
                    return;
                }
                _selectedTestSuite = value;
                if (_selectedTestSuite != null)
                {
                    StoredSelectedTestSuite = _selectedTestSuite.Title;
                }
                OnPropertyChanged(nameof(SelectedTestSuite));
                if (CreateReportCommand != null)
                {
                    CreateReportCommand.CallEventCanExecuteChanged();
                }
            }
        }

        /// <summary>
        /// Gets the list of all available test configurations
        /// </summary>
        public IList<ITfsTestConfiguration> TestConfigurations
        {
            get
            {
                var serverBuild = SelectedServerBuild;
                if (serverBuild != null && serverBuild.BuildNumber == AllServerBuilds.AllServerBuildsId)
                // null means use all server builds
                {
                    serverBuild = null;
                }

                if (_lastUsedSelectedServerBuildForTestConfigurations != serverBuild
                    || _lastUsedTestPlanForTestConfigurations != SelectedTestPlan)
                {
                    _lastUsedTestPlanForTestConfigurations = SelectedTestPlan;
                    _testConfigurations = null;
                    if (SelectedTestPlan == null)
                    {
                        return _testConfigurations;
                    }

                    _lastUsedSelectedServerBuildForTestConfigurations = serverBuild;

                    _testConfigurations = new List<ITfsTestConfiguration> { new AllTestConfigurations() };
                    try
                    {
                        foreach (var element in TestAdapter.GetTestConfigurationsForTestPlanWithTestResults(serverBuild, SelectedTestPlan))
                            _testConfigurations.Add(element);
                    }
                    finally
                    {
                        OnPropertyChanged(nameof(TestConfigurations));
                        if (ViewDispatcher != null)
                        {
                            ViewDispatcher.BeginInvoke(new Action(() =>
                            {
                                var testConfigurationToSelect = _testConfigurations[0];
                                if (!string.IsNullOrEmpty(StoredSelectedTestConfiguration) && _testConfigurations.Any(conf => conf.Name == StoredSelectedTestConfiguration))
                                {
                                    testConfigurationToSelect = _testConfigurations.First(conf => conf.Name == StoredSelectedTestConfiguration);
                                }
                                SelectedTestConfiguration = testConfigurationToSelect;
                            }));
                        }
                    }
                }
                return _testConfigurations;
            }
        }

        /// <summary>
        /// Gets or sets selected test configuration.
        /// </summary>
        public ITfsTestConfiguration SelectedTestConfiguration
        {
            get
            {
                return _selectedTestConfiguration;
            }
            set
            {
                if (_selectedTestConfiguration == value)
                    return;

                // on first value assignment use default value
                if (_selectedTestConfiguration == null && ViewDispatcher != null && ViewDispatcher.IsDispatching)
                {
                    var defaultTestConfig = TestConfigurations.FirstOrDefault(x => x.Name == _defaultConfigurationName);
                    value = defaultTestConfig ?? value;
                }

                _selectedTestConfiguration = value;
                if (_selectedTestConfiguration != null)
                    StoredSelectedTestConfiguration = _selectedTestConfiguration.Name;
                OnPropertyChanged(nameof(SelectedTestConfiguration));
                if (ViewDispatcher != null)
                {
                    ViewDispatcher.Invoke(() => CreateReportCommand.CallEventCanExecuteChanged());
                }
            }
        }

        /// <summary>
        /// Gets or sets the information, whether the document structure should be created.
        /// </summary>
        public bool CreateDocumentStructure
        {
            get { return _createDocumentStructure; }
            set
            {
                if (_createDocumentStructure == value)
                {
                    return;
                }
                _createDocumentStructure = value;
                StoredCreateDocumentStructure = _createDocumentStructure;
                OnPropertyChanged(nameof(CreateDocumentStructure));
                OnPropertyChanged(nameof(IsSkipLevelsEnabled));
            }
        }

        /// <summary>
        /// Gets all possible / supported document structures
        /// </summary>
        public static IList<DocumentStructureType> DocumentStructureTypes
        {
            get
            {
                return Enum.GetValues(typeof(DocumentStructureType)).Cast<DocumentStructureType>().ToList();
            }
        }

        /// <summary>
        /// Gets all possible / supported positions for the configuration of the tests
        /// </summary>
        public static IList<ConfigurationPositionType> ConfigurationPositionType
        {
            get
            {
                return Enum.GetValues(typeof(ConfigurationPositionType)).Cast<ConfigurationPositionType>().ToList();
            }
        }

        /// <summary>
        /// Gets or sets <see cref="DocumentStructureType"/>.
        /// </summary>
        public DocumentStructureType SelectedDocumentStructureType
        {
            get
            {
                return _selectedDocumentStructureType;
            }
            set
            {
                if (_selectedDocumentStructureType == value)
                {
                    return;
                }
                _selectedDocumentStructureType = value;
                StoredSelectedDocumentStructureType = _selectedDocumentStructureType;
                OnPropertyChanged(nameof(SelectedDocumentStructureType));
                OnPropertyChanged(nameof(IsSkipLevelsEnabled));
            }
        }

        /// <summary>
        /// Gets or sets <see cref="ConfigurationPositionType"/>.
        /// </summary>
        public ConfigurationPositionType SelectedConfigurationPositionType
        {
            get
            {
                return _selectedConfigurationPositionType;
            }
            set
            {
                if (_selectedConfigurationPositionType == value)
                {
                    return;
                }
                _selectedConfigurationPositionType = value;
                StoredSelectedConfigurationPositionType = _selectedConfigurationPositionType;
                OnPropertyChanged(nameof(SelectedConfigurationPositionType));
            }
        }



        /// <summary>
        /// Gets or sets count of levels to ignore.
        /// Value is used only for <see cref="DocumentStructureType.IterationPath"/> and <see cref="DocumentStructureType.AreaPath"/>.
        /// </summary>
        public int SkipLevels
        {
            get { return _skipLevels; }
            set
            {
                if (_skipLevels == value)
                {
                    return;
                }
                _skipLevels = value;
                StoredSkipLevels = _skipLevels;
                OnPropertyChanged(nameof(SkipLevels));
            }
        }

        /// <summary>
        /// Gets the flag whether defines if the SkipLevels control is enabled.
        /// </summary>
        public bool IsSkipLevelsEnabled
        {
            get
            {
                return !(!CreateDocumentStructure || SelectedDocumentStructureType == DocumentStructureType.TestPlanHierarchy);
            }
        }


        /// <summary>
        /// Gets the flag whether defines if the SkipLevels control is enabled.
        /// </summary>
        public bool IsMostRecentAllconfigurationsEnabled
        {
            get
            {
                return IncludeOnlyMostRecentTestResult;
            }
        }


        /// <summary>
        /// Gets or sets whether the test configurations are to include in generated document.
        /// </summary>
        public bool IncludeTestConfigurations
        {
            get { return _includeTestConfigurations; }
            set
            {
                if (_includeTestConfigurations == value)
                {
                    return;
                }
                _includeTestConfigurations = value;
                StoredIncludeTestConfigurations = _includeTestConfigurations;
                OnPropertyChanged(nameof(IncludeTestConfigurations));
            }
        }

        /// <summary>
        /// Gets all possible / supported test case sorts
        /// </summary>
        public static IList<TestCaseSortType> TestCaseSorts
        {
            get
            {
                return Enum.GetValues(typeof(TestCaseSortType)).Cast<TestCaseSortType>().ToList();
            }
        }

        /// <summary>
        /// Gets or sets <see cref="TestCaseSortType"/>.
        /// </summary>
        public TestCaseSortType SelectedTestCaseSortType
        {
            get
            {
                return _selectedTestCaseSortType;
            }
            set
            {
                if (_selectedTestCaseSortType == value)
                {
                    return;
                }
                _selectedTestCaseSortType = value;
                StoredSelectedTestCaseSortType = _selectedTestCaseSortType;
                OnPropertyChanged(nameof(SelectedTestCaseSortType));
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to include only most recent test result.
        /// </summary>
        /// <value>
        /// <c>true</c> if only the latest test result for a test case is reported otherwise, <c>false</c>.
        /// </value>
        public bool IncludeOnlyMostRecentTestResult
        {
            get
            {
                return _includeOnlyMostRecentTestResult;
            }
            set
            {
                if (_includeOnlyMostRecentTestResult == value) return;
                _includeOnlyMostRecentTestResult = value;
                StoredIncludeOnlyMostRecentTestResult = _includeOnlyMostRecentTestResult;

                OnPropertyChanged(nameof(IncludeOnlyMostRecentTestResult));
                OnPropertyChanged(nameof(IsMostRecentAllconfigurationsEnabled));
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to the last test results for all configurations
        /// </summary>
        /// <value>
        /// <c>true</c> if only the latest test result for a test case is reported otherwise, <c>false</c>.
        /// </value>
        public bool IncludeOnlyMostRecentTestResultForAllConfigurations
        {
            get
            {
                return _includeOnlyMostRecentTestResultForAllConfigurations;
            }
            set
            {
                if (_includeOnlyMostRecentTestResultForAllConfigurations == value)
                {
                    return;
                }
                _includeOnlyMostRecentTestResultForAllConfigurations = value;
                StoredIncludeOnlyMostRecentTestResultForAllConfigurations = _includeOnlyMostRecentTestResultForAllConfigurations;

                OnPropertyChanged(nameof(IncludeOnlyMostRecentTestResultForAllConfigurations));
            }
        }

        #endregion Public binding properties

        #region Private methods
        private void SetTestReportDefaults(ITestResultReportDefault defaultValue)
        {
            if (defaultValue == null)
                return;

            //directly initialized
            CreateDocumentStructure = defaultValue.CreateDocumentStructure;
            SelectedDocumentStructureType = defaultValue.DocumentStructureType;
            IncludeTestConfigurations = defaultValue.IncludeTestConfigurations;
            SelectedConfigurationPositionType = defaultValue.ConfigurationPositionType;
            SelectedTestCaseSortType = defaultValue.TestCaseSortType;
            SkipLevels = defaultValue.SkipLevels;
            IncludeOnlyMostRecentTestResult = defaultValue.IncludeMostRecentTestResult;

            //logical combinations
            if (IncludeOnlyMostRecentTestResult) IncludeOnlyMostRecentTestResultForAllConfigurations = defaultValue.IncludeMostRecentTestResultForAllSelectedConfigurations;

            if (TestConfigurations != null)
            {
                var tesconfig = TestConfigurations.FirstOrDefault(x => x.Name == defaultValue.SelectTestConfiguration);
                SelectedTestConfiguration = tesconfig ?? SelectedTestConfiguration;
            }

            // lazy initialized
            _defaultTestPlanName = defaultValue.SelectTestPlan;
            _defaultConfigurationName = defaultValue.SelectTestConfiguration;
            _defaultServerBuildName = defaultValue.SelectBuild;
        }

        #endregion Private methods

        /// <summary>
        /// Gets the list of all available server builds
        /// Note: Property ServerBuilds does the same but asynchronous - this is not applicable for the console extension
        /// </summary>
        /// <returns></returns>
        public IList<ITfsServerBuild> GetServerBuilds()
        {
            if (_serverBuilds == null)
            {
                _serverBuilds = new List<ITfsServerBuild> { new AllServerBuilds() };
                try
                {
                    _serverBuilds = _serverBuilds.Concat(TestAdapter.GetAvailableServerBuilds()).ToList();
                }
                finally
                {
                    OnPropertyChanged(nameof(ServerBuilds));
                }
            }
            return _serverBuilds;
        }
        #region Private command methods

        /// <summary>
        /// Execute method for 'Create Report'.
        /// </summary>
        /// <param name="parameter">Parameter for command.</param>
        private void ExecuteCreateReportCommand(object parameter)
        {
            var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
            if (infoStorage != null)
            {
                infoStorage.ClearAll();
            }

            StartBackgroundWorker(true, CreateReport);
        }

        /// <summary>
        /// Can execute method for 'Create Report'.
        /// </summary>
        /// <param name="parameter">Parameter for command.</param>
        /// <returns>True if the command can be executed. Otherwise false.</returns>
        private bool CanExecuteCreateReportCommand(object parameter)
        {
            return !(SelectedServerBuild == null || SelectedTestConfiguration == null || SelectedTestPlan == null);
        }

        #endregion Private command methods

        #region Private create report methods

        /// <summary>
        /// The method creates report for selected test suite.
        /// </summary>
        public void CreateReport()
        {
            SyncServiceTrace.D(Resources.CreateReportInfo);

            SyncServiceDocumentModel.TestReportRunning = true;
            try
            {

                var testConfiguration = SelectedTestConfiguration;
                if (testConfiguration.Id == AllTestConfigurations.AllTestConfigurationsId) // null means use all configurations.
                {
                    testConfiguration = null;
                }
                var serverBuild = SelectedServerBuild;
                if (serverBuild.BuildNumber == AllServerBuilds.AllServerBuildsId) // null means use all server builds
                {
                    serverBuild = null;
                }

                ProcessPreOperations();

                // Common part
                if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.AboveTestPlan)
                {
                    CreateConfigurationPart();
                }
                // Document part
                if (CreateDocumentStructure)
                {
                    // Structured document
                    switch (SelectedDocumentStructureType)
                    {
                        case DocumentStructureType.AreaPath:
                            CreateReportByAreaPath(SelectedTestPlan, SelectedTestSuite, SelectedTestConfiguration, serverBuild);
                            break;
                        case DocumentStructureType.IterationPath:
                            CreateReportByIterationPath(SelectedTestPlan, SelectedTestSuite, testConfiguration, serverBuild);
                            break;
                        case DocumentStructureType.TestPlanHierarchy:
                            CreateReportByTestPlanHierarchy(SelectedTestPlan, SelectedTestSuite, testConfiguration, serverBuild);
                            break;
                    }
                }
                else
                    // Unstructured document
                    CreateReportUnstructured(SelectedTestPlan, SelectedTestSuite, testConfiguration, serverBuild);

                var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
                var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
                var testReportHelper = new TestReportHelper(TestAdapter, testReport, config, CancellationPending);

                if (TestReportingProgressCancellationService.CheckIfContinue())
                {
                    testReportHelper.CreateSummaryPage(WordDocument, SelectedTestPlan);
                    ProcessPostOperations();
                }
                else
                {
                    SyncServiceTrace.I(Resources.DoImportError);
                }

                if (!CancellationPending())
                {
                    // Set the 'Report generated' only if the report was not canceled
                    SyncServiceDocumentModel.TestReportGenerated = true;
                    if (ViewDispatcher != null)
                    {
                        ViewDispatcher.Invoke(() => CreateReportCommand.CallEventCanExecuteChanged());
                    }
                    StoreReportData();
                }
            }
            catch (Exception ex)
            {
                SyncServiceTrace.LogException(ex);
                var infoStorageService = SyncServiceFactory.GetService<IInfoStorageService>();
                if (infoStorageService == null) throw;
                IUserInformation info = new UserInformation
                {
                    Text = Resources.TestResult_Error,
                    Explanation = ex is OperationCanceledException ? ex.Message : ex.ToString(),
                    Type = UserInformationType.Error
                };
                infoStorageService.AddItem(info);
            }
            finally
            {
                SyncServiceDocumentModel.TestReportRunning = false;
            }
        }

        private IWord2007TestReportAdapter testReportOperation;
        /// <summary>
        /// Process an Operation
        /// </summary>
        private void ProcessPreOperations()
        {
            //Process Pre Operations
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            testReportOperation = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Create Report helper
            var operation = config.ConfigurationTest.GetPreOperationsForTestResult();
            var testReportHelper = new TestReportHelper(TestAdapter, testReportOperation, config, CancellationPending);

            testReportOperation.PrepareDocumentForLongTermOperation();
            testReportHelper.ProcessOperations(operation);

        }

        /// <summary>
        /// Process an Operation
        /// </summary>
        private void ProcessPostOperations()
        {
            //Process Pre Operations
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            // Create Report helper
            var operation = config.ConfigurationTest.GetPostOperationsForTestResult();
            var testReportHelper = new TestReportHelper(TestAdapter, testReportOperation, config, CancellationPending);

            testReportHelper.ProcessOperations(operation);
            testReportOperation.UndoPreparationsDocumentForLongTermOperation();
        }

        /// <summary>
        /// The method writes out the configuration information.
        /// </summary>
        private void CreateConfigurationPart()
        {
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);

            // Create report helper
            var testReportHelper = new TestReportHelper(TestAdapter, testReport, config, CancellationPending);

            // Insert header of test configurations
            var templateName = config.ConfigurationTest.ConfigurationTestResult.TestConfigurationElementTemplate;
            testReportHelper.InsertHeaderTemplate(config.ConfigurationTest.GetHeaderTemplate(templateName));

            // Determine configurations
            var configList = new List<ITfsTestConfiguration>();
            if (SelectedTestConfiguration.Id == AllTestConfigurations.AllTestConfigurationsId)
            {
                foreach (var testConfiguration in TestConfigurations)
                {
                    if (testConfiguration.Id == AllTestConfigurations.AllTestConfigurationsId)
                    {
                        continue;
                    }
                    configList.Add(testConfiguration);
                }
            }
            else
            {
                configList.Add(SelectedTestConfiguration);
            }
            // Iterate through configurations
            foreach (var testConfiguration in configList)
            {
                testReportHelper.InsertTestConfiguration(templateName, TestAdapter.GetTestConfigurationDetail(testConfiguration));
            }
        }

        /// <summary>
        /// The method creates unstructured report - writes out all test cases in one block.
        /// </summary>
        /// <param name="testPlan"><see cref="ITfsTestPlan"/> is the source of test cases for to write out with its test results.</param>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> is the source of test cases for to write out with its test results.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        /// If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        /// If <c>null</c>, the build is no matter.</param>
        private void CreateReportUnstructured(ITfsTestPlan testPlan, ITfsTestSuite testSuite, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild)
        {
            if (testPlan == null || testSuite == null)
            {
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Get the detailed test plan.
            var testPlanDetail = TestAdapter.GetTestPlanDetail(testPlan);
            if (testPlanDetail == null)
            {
                return;
            }
            // Get the detailed test suite.
            var testSuiteDetail = TestAdapter.GetTestSuiteDetail(testSuite);
            if (testSuiteDetail == null)
            {
                return;
            }

            // Create report helper
            var testReportHelper = new TestReportHelper(TestAdapter, testReport, config, CancellationPending);
            testReportHelper.IncludeMostRecentTestResults = IncludeOnlyMostRecentTestResult;
            testReportHelper.IncludeOnlyMostRecentTestResultForAllConfigurations = IncludeOnlyMostRecentTestResultForAllConfigurations;
            testReportHelper.CurrentTestConfiguration = SelectedTestConfiguration;

            // Insert test plan template
            var templateName = config.ConfigurationTest.ConfigurationTestResult.TestPlanTemplate;
            testReportHelper.InsertTestPlanTemplate(templateName, testPlanDetail);
            // Common part
            if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathTestPlan)
            {
                CreateConfigurationPart();
            }
            // Insert root test suite template - if null, root test suite template = test suite template
            templateName = config.ConfigurationTest.ConfigurationTestResult.RootTestSuiteTemplate;
            testReportHelper.InsertTestSuiteTemplate(templateName, testSuiteDetail);

            var expandSharedSteps = config.ConfigurationTest.ExpandSharedSteps;
            // Get test cases of the test suite - with test cases in sub test suites
            var testCases = TestAdapter.GetAllTestCases(testSuite, expandSharedSteps);

            // TestCasesHelper need document structure, but the enumerations has not value 'None'
            // We will use the functionality without this structure capability
            var helper = new TestCaseHelper(TestAdapter, testPlan, testConfiguration, serverBuild, testCases, SelectedDocumentStructureType);
            // Get sorted test cases
            var sortedTestCases = helper.GetTestCases(SelectedTestCaseSortType);
            // Test if test cases exists
            if (sortedTestCases != null && sortedTestCases.Count > 0)
            {
                PrintOutTestResults(sortedTestCases, config, testPlan, testConfiguration, serverBuild, testReportHelper, testSuite);
            }
        }

        /// <summary>
        /// The method creates report structured by area path.
        /// </summary>
        /// <param name="testPlan"><see cref="ITfsTestPlan"/> is the source of test cases for to write out with its test results.</param>
        /// /// <param name="testSuite"><see cref="ITfsTestSuite"/> is the source of test cases for to write out with its test results.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        /// If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        /// If <c>null</c>, the build is no matter.</param>
        private void CreateReportByAreaPath(ITfsTestPlan testPlan, ITfsTestSuite testSuite, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild)
        {
            SyncServiceTrace.D(Resources.CreateReportByAreaPath);
            if (testPlan == null || testSuite == null)
            {
                SyncServiceTrace.D(Resources.TestPlanOrSuiteNotExist);
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Get the detailed test plan.
            var testPlanDetail = TestAdapter.GetTestPlanDetail(testPlan);
            if (testPlanDetail == null)
            {
                return;
            }
            // Get the detailed test suite.
            var testSuiteDetail = TestAdapter.GetTestSuiteDetail(testSuite);
            if (testSuiteDetail == null)
            {
                return;
            }

            // Create report helper
            var testReportHelper = new TestReportHelper(TestAdapter, testReport, config, CancellationPending);
            testReportHelper.IncludeMostRecentTestResults = IncludeOnlyMostRecentTestResult;
            testReportHelper.IncludeOnlyMostRecentTestResultForAllConfigurations = IncludeOnlyMostRecentTestResultForAllConfigurations;
            testReportHelper.CurrentTestConfiguration = SelectedTestConfiguration;

            // Insert test plan template
            var templateName = config.ConfigurationTest.ConfigurationTestResult.TestPlanTemplate;
            testReportHelper.InsertTestPlanTemplate(templateName, testPlanDetail);
            // Insert root test suite template: if null, root test suite template = test suite template
            templateName = config.ConfigurationTest.ConfigurationTestResult.RootTestSuiteTemplate;
            testReportHelper.InsertTestSuiteTemplate(templateName, testSuiteDetail);

            // Get test cases of the test suite - with test cases in sub test suites
            var expandSharedSteps = config.ConfigurationTest.ExpandSharedSteps;
            var testCases = TestAdapter.GetAllTestCases(testSuite, expandSharedSteps);

            SyncServiceTrace.D(Resources.NumberOfAllTestCases + testCases.Count);
            // Create test case helper
            var helper = new TestCaseHelper(TestAdapter, testPlan, testConfiguration, serverBuild, testCases, SelectedDocumentStructureType);
            // Get all path elements from all test cases
            var pathElements = helper.GetPathElements(SkipLevels);

            // Iterate through path elements
            foreach (var pathElement in pathElements)
            {
                if (CancellationPending()) return;

                // Insert heading
                testReportHelper.InsertHeadingText(pathElement.PathPart, pathElement.Level - SkipLevels);
                // Get sorted test cases
                var sortedTestCases = helper.GetTestCases(pathElement, SelectedTestCaseSortType);
                if (sortedTestCases != null && sortedTestCases.Count > 0)
                {
                    PrintOutTestResults(sortedTestCases, config, testPlan, testConfiguration, serverBuild, testReportHelper, null);
                }
            }
            // Common part
            if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathTestPlan)
            {
                CreateConfigurationPart();
            }
        }

        /// <summary>
        /// The method creates report structured by iteration path.
        /// </summary>
        /// <param name="testPlan"><see cref="ITfsTestPlan"/> is the source of test cases for to write out with its test results.</param>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> is the source of test cases for to write out with its test results.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        /// If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        /// If <c>null</c>, the build is no matter.</param>
        private void CreateReportByIterationPath(ITfsTestPlan testPlan, ITfsTestSuite testSuite, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild)
        {
            SyncServiceTrace.D(Resources.CreateReportByIterationPath);
            if (testPlan == null || testSuite == null)
            {
                SyncServiceTrace.D(Resources.TestPlanOrSuiteNotExist);
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Get the detailed test plan.
            var testPlanDetail = TestAdapter.GetTestPlanDetail(testPlan.RootTestSuite);
            if (testPlanDetail == null)
            {
                return;
            }
            // Get the detailed test suite.
            var testSuiteDetail = TestAdapter.GetTestSuiteDetail(testSuite);
            if (testSuiteDetail == null)
            {
                return;
            }

            // Create report helper
            var testReportHelper = new TestReportHelper(TestAdapter, testReport, config, CancellationPending);
            testReportHelper.IncludeMostRecentTestResults = IncludeOnlyMostRecentTestResult;
            testReportHelper.IncludeOnlyMostRecentTestResultForAllConfigurations = IncludeOnlyMostRecentTestResultForAllConfigurations;
            testReportHelper.CurrentTestConfiguration = SelectedTestConfiguration;

            // Insert test plan template
            var templateName = config.ConfigurationTest.ConfigurationTestResult.TestPlanTemplate;
            testReportHelper.InsertTestPlanTemplate(templateName, testPlanDetail);
            // Insert root test suite template: if null, root test suite template = test suite template
            templateName = config.ConfigurationTest.ConfigurationTestResult.RootTestSuiteTemplate;
            testReportHelper.InsertTestSuiteTemplate(templateName, testSuiteDetail);


            // Get test cases of the test suite - with test cases in sub test suites
            var expandSharedSteps = config.ConfigurationTest.ExpandSharedSteps;
            var testCases = TestAdapter.GetAllTestCases(testSuiteDetail.TestSuite, expandSharedSteps);
            SyncServiceTrace.D(Resources.NumberOfAllTestCases + testCases.Count);
            // Create test case helper
            var helper = new TestCaseHelper(TestAdapter, testPlan, testConfiguration, serverBuild, testCases, SelectedDocumentStructureType);
            // Get all path elements from all test cases
            var pathElements = helper.GetPathElements(SkipLevels);

            // Iterate through path elements
            foreach (var pathElement in pathElements)
            {
                if (CancellationPending()) return;

                // Insert heading
                testReportHelper.InsertHeadingText(pathElement.PathPart, pathElement.Level - SkipLevels);
                // Get sorted test cases
                var sortedTestCases = helper.GetTestCases(pathElement, SelectedTestCaseSortType);
                if (sortedTestCases != null && sortedTestCases.Count > 0)
                {
                    PrintOutTestResults(sortedTestCases, config, testPlan, testConfiguration, serverBuild, testReportHelper, null);
                }
            }

            // Remove the inserted bookmarks
            testReport.RemoveBookmarks();

            // Common part
            if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathTestPlan)
            {
                CreateConfigurationPart();
            }
        }

        /// <summary>
        /// The method generates report by test plan structure.
        /// </summary>
        /// <param name="testPlan"><see cref="ITfsTestPlan"/> is the source of test cases for to write out with its test results.</param>
        /// <param name="testSuite"><see cref="ITfsTestPlan"/> is the source of test cases for to write out with its test results.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        /// If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        /// If <c>null</c>, the build is no matter.</param>
        private void CreateReportByTestPlanHierarchy(ITfsTestPlan testPlan, ITfsTestSuite testSuite, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild)
        {
            SyncServiceTrace.D(Resources.CreateReportByTestPlanHierarchy);
            if (testPlan == null || testSuite == null)
            {
                SyncServiceTrace.D(Resources.TestPlanOrSuiteNotExist);
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Get the detailed test plan.
            var testPlanDetail = TestAdapter.GetTestPlanDetail(testPlan);
            if (testPlanDetail == null)
            {
                return;
            }
            // Create report helper
            var testReportHelper = new TestReportHelper(TestAdapter, testReport, config, CancellationPending);
            testReportHelper.IncludeMostRecentTestResults = IncludeOnlyMostRecentTestResult;
            testReportHelper.IncludeOnlyMostRecentTestResultForAllConfigurations = IncludeOnlyMostRecentTestResultForAllConfigurations;
            testReportHelper.CurrentTestConfiguration = SelectedTestConfiguration;

            // Insert test plan template
            var templateName = config.ConfigurationTest.ConfigurationTestResult.TestPlanTemplate;
            testReportHelper.InsertTestPlanTemplate(templateName, testPlanDetail);
            // Common part
            if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathTestPlan)
            {
                CreateConfigurationPart();
            }
            // Call recursion
            CreateReportByTestSuiteHierarchy(testPlan, testSuite, 1, testConfiguration, serverBuild, true);

        }

        /// <summary>
        /// The method generates report by test suite structure - recursion method called from CreateReportByTestPlanHierarchy.
        /// </summary>
        /// <param name="testPlan"><see cref="ITfsTestPlan"/> is the source of test cases to write out.</param>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> is the source of test cases to write out.</param>
        /// <param name="headingLevel">Level of the heading. First level is 1.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        /// If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        /// If <c>null</c>, the build is no matter.</param>
        /// <param name="firstSuite">The first test suite should be used</param>
        private void CreateReportByTestSuiteHierarchy(ITfsTestPlan testPlan, ITfsTestSuite testSuite, int headingLevel, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild, bool firstSuite)
        {
            if (headingLevel > 10)
            {
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            // Get the detailed test suite.
            var testSuiteDetail = TestAdapter.GetTestSuiteDetail(testSuite);
            if (testSuiteDetail == null)
            {
                return;
            }

            // Check if this test suite or any child test suite contains at least one test case.
            // If no, to nothing.
            var expandSharedSteps = config.ConfigurationTest.ExpandSharedSteps;
            var allTestCases = TestAdapter.GetAllTestCases(testSuite, expandSharedSteps);
            SyncServiceTrace.D(Resources.NumberOfAllTestCases + allTestCases.Count);
            if (allTestCases == null || allTestCases.Count == 0)
            {
                return;
            }


            IWord2007TestReportAdapter testReport = null;
            if (config.AttachmentFolderMode == AttachmentFolderMode.BasedOnTestSuite)
            {
                testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config, testSuite, allTestCases);
            }
            else
            {
                testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            }

            //If there is no test that has a result, it was immediately out of the methods
            bool includeTestCasesWithoutResults = config.ConfigurationTest.ConfigurationTestResult.IncludeTestCasesWithoutResults;
            if (!includeTestCasesWithoutResults && !TestAdapter.TestResultExists(testPlan, testSuite, allTestCases, testConfiguration, serverBuild))
            {
                return;
            }

            // Create report helper
            var testReportHelper = new TestReportHelper(TestAdapter, testReport, config, CancellationPending);
            testReportHelper.IncludeMostRecentTestResults = IncludeOnlyMostRecentTestResult;
            testReportHelper.IncludeOnlyMostRecentTestResultForAllConfigurations = IncludeOnlyMostRecentTestResultForAllConfigurations;
            testReportHelper.CurrentTestConfiguration = SelectedTestConfiguration;

            // Insert heading
            testReportHelper.InsertHeadingText(testSuiteDetail.TestSuite.Title, headingLevel);
            var templateName = config.ConfigurationTest.ConfigurationTestResult.TestSuiteTemplate;
            // Insert leaf test suite template for leaf suites
            if (testSuiteDetail.TestSuite.TestSuites == null || !testSuiteDetail.TestSuite.TestSuites.Any())
            {
                templateName = config.ConfigurationTest.ConfigurationTestResult.LeafTestSuiteTemplate;

            }
            // Insert root test suite template: if null, root test suite template = test suite template
            if (firstSuite)
            {
                templateName = config.ConfigurationTest.ConfigurationTestResult.RootTestSuiteTemplate;
            }
            testReportHelper.InsertTestSuiteTemplate(templateName, testSuiteDetail);

            // Print the test configuration beneath the first testsuite
            if (!_testConfigurationInformationPrinted && IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathFirstTestSuite)
            {
                CreateConfigurationPart();
                _testConfigurationInformationPrinted = true;
            }

            // Get test cases of the test suite - only test cases in this test suite
            var testCases = TestAdapter.GetTestCases(testSuite, expandSharedSteps);
            SyncServiceTrace.D(Resources.NumberOfTestCasesBasedOnSuite + testCases.Count);
            // TestCasesHelper need document structure, but the enumerations has not value 'None'
            // We will use the functionality without this structure capability
            var helper = new TestCaseHelper(TestAdapter, testPlan, testConfiguration, serverBuild, testCases, SelectedDocumentStructureType);
            // Get sorted test cases
            var sortedTestCases = helper.GetTestCases(SelectedTestCaseSortType);
            // Test if test cases exists
            if (sortedTestCases != null && sortedTestCases.Count > 0)
            {
                PrintOutTestResults(sortedTestCases, config, testPlan, testConfiguration, serverBuild, testReportHelper, testSuite);
            }

            // Check child suites
            if (testSuiteDetail.TestSuite.TestSuites != null)
            {
                // Iterate through child test suites
                foreach (var childTestSuite in testSuiteDetail.TestSuite.TestSuites)
                {
                    if (CancellationPending())
                        return;
                    CreateReportByTestSuiteHierarchy(testPlan, childTestSuite, headingLevel + 1, testConfiguration, serverBuild, false);
                }
            }

            //Print the configuration information beneath all TestSuites
            if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathTestSuites)
            {
                CreateConfigurationPart();
            }

        }

        private void PrintOutTestResults(IEnumerable<ITfsTestCaseDetail> sortedTestCases, IConfiguration config, ITfsTestPlan testPlan, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild, TestReportHelper testReportHelper, ITfsTestSuite testSuite)
        {
            SyncServiceTrace.D(Resources.PrintTestResults);

            var includeTestCasesWithoutResults = config.ConfigurationTest.ConfigurationTestResult.IncludeTestCasesWithoutResults;
            // Iterate all test cases
            foreach (var testCase in sortedTestCases)
            {
                if (CancellationPending())
                {
                    return;
                }
                if (!includeTestCasesWithoutResults && !TestAdapter.TestResultExists(testPlan, testSuite, testCase, testConfiguration, serverBuild))
                {
                    continue;
                }

                // Report the test case
                var templateName = config.ConfigurationTest.ConfigurationTestResult.TestCaseElementTemplate;
                testReportHelper.InsertTestCase(templateName, testCase);


                // Get relevant test results for test case / plan / config

                IList<ITfsTestResultDetail> testResults;
                if (IncludeOnlyMostRecentTestResult)
                {
                    //Get the last testresult for each configuration
                    if (IncludeOnlyMostRecentTestResultForAllConfigurations)
                    {
                        testResults = TestAdapter.GetLatestTestResultsForAllSelectedConfigurations(testPlan, testCase.TestCase, testConfiguration, serverBuild, testSuite, false);
                    }
                    else
                    {
                        testResults = TestAdapter.GetLatestTestResults(testPlan, testCase.TestCase, testConfiguration, serverBuild, testSuite, false);
                    }
                }
                else
                {
                    Dictionary<int, int> lastTestRunPerConfig;
                    testResults = TestAdapter.GetTestResults(testPlan, testCase.TestCase, testConfiguration, serverBuild, testSuite, false, out lastTestRunPerConfig);
                }

                if (testResults.Count() != 0)
                {
                    // Report the test result header
                    templateName = config.ConfigurationTest.ConfigurationTestResult.TestResultElementTemplate;
                    testReportHelper.InsertHeaderTemplate(config.ConfigurationTest.GetHeaderTemplate(templateName));

                    // Report the test results
                    foreach (var testResult in testResults)
                    {
                        if (CancellationPending()) return;

                        testReportHelper.InsertTestResult(templateName, testResult);

                    }
                }


            }
        }

        #endregion Private create report methods
    }
}
