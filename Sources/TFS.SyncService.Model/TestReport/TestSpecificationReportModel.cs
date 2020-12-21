#region Usings
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Contracts.Word;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Model.Properties;
using AIT.TFS.SyncService.Service.InfoStorage;
using Microsoft.Office.Interop.Word;
using AIT.TFS.SyncService.Contracts.Adapter;
#endregion

namespace AIT.TFS.SyncService.Model.TestReport
{
    /// <summary>
    /// Model class for TestSpecificationReport.
    /// </summary>
    public sealed class TestSpecificationReportModel : TestReportModel
    {
        #region Private fields

        private readonly ObservableCollection<ITfsTestSuite> _treeViewTestSuites = new ObservableCollection<ITfsTestSuite>();
        private IList<ITfsTestPlan> _testPlans;
        private ITfsTestPlan _selectedTestPlan;
        private IList<ITfsTestConfiguration> _testConfigurations;
        private ITfsTestConfiguration _selectedTestConfiguration;
        private bool _includeTestConfigurations;
        private object _selectedTreeViewItem;
        private ITfsTestSuite _selectedTestSuite;
        private bool _createDocumentStructure;
        private DocumentStructureType _selectedDocumentStructureType;
        private ConfigurationPositionType _selectedConfigurationPositionType;
        private int _skipLevels;
        private TestCaseSortType _selectedTestCaseSortType;
        private bool _testConfigurationInformationPrinted;

        private string _defaultTestPlanName;
        private string _defaultTestSuiteName;

        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TestSpecificationReportModel"/> class.
        /// </summary>
        /// <param name="syncServiceDocumentModel">The <see cref="ISyncServiceModel"/> to obtain document settings.</param>
        /// <param name="dispatcher">Dispatcher of associated view.</param>
        /// <param name="testAdapter"><see cref="ITfsTestAdapter"/> to examine test information.</param>
        /// <param name="wordRibbon">Interface to the word ribbon</param>
        /// <param name="testReportingProgressCancellationService">The progess cancellation service used to check at certain points if a cancellation has been triggered and further steps should be skipped.</param>
        public TestSpecificationReportModel(ISyncServiceDocumentModel syncServiceDocumentModel, IViewDispatcher dispatcher,
            ITfsTestAdapter testAdapter, IWordRibbon wordRibbon, ITestReportingProgressCancellationService testReportingProgressCancellationService)
            : base(syncServiceDocumentModel, dispatcher, testAdapter, wordRibbon, testReportingProgressCancellationService)
        {
            CreateDocumentStructure = StoredCreateDocumentStructure;
            SkipLevels = StoredSkipLevels;
            SelectedDocumentStructureType = StoredSelectedDocumentStructureType;
            SelectedTestCaseSortType = StoredSelectedTestCaseSortType;
            CreateReportCommand = new ViewCommand(ExecuteCreateReportCommand, CanExecuteCreateReportCommand);
            SetTestReportDefaults(syncServiceDocumentModel.Configuration.ConfigurationTest.ConfigurationTestSpecification.DefaultValues);

            WordDocument = syncServiceDocumentModel.WordDocument as Document;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TestSpecificationReportModel"/> class for the console extension
        /// </summary>
        /// <param name="syncServiceDocumentModel">The <see cref="ISyncServiceModel"/> to obtain document settings.</param>
        /// <param name="testAdapter"><see cref="ITfsTestAdapter"/> to examine test information.</param>
        /// <param name="sort"></param>
        /// <param name="testReportingProgressCancellationService">The progess cancellation service used to check at certain points if a cancellation has been triggered and further steps should be skipped.</param>
        /// <param name="plan">The selected <see cref="ITfsTestPlan"/> used for the report generation.</param>
        /// <param name="suite">The selected <see cref="ITfsTestSuite"/> used for the report generation.</param>
        /// <param name="documentStructure">The information if a document structure should be created.</param>
        /// <param name="skipLevels">The level count to ignore generation.</param>
        /// <param name="structureType"></param>
        public TestSpecificationReportModel(ISyncServiceDocumentModel syncServiceDocumentModel, ITfsTestAdapter testAdapter, ITfsTestPlan plan, ITfsTestSuite suite, bool documentStructure, int skipLevels, DocumentStructureType structureType, TestCaseSortType sort, ITestReportingProgressCancellationService testReportingProgressCancellationService)
            : base(syncServiceDocumentModel, testAdapter, testReportingProgressCancellationService)
        {
            SelectedTestPlan = plan;
            SelectedTestSuite = suite;
            CreateDocumentStructure = documentStructure;
            SkipLevels = skipLevels;
            SelectedDocumentStructureType = structureType;
            SelectedTestCaseSortType = sort;
            CreateReportCommand = new ViewCommand(ExecuteCreateReportCommand, CanExecuteCreateReportCommand);
            WordDocument = syncServiceDocumentModel.WordDocument as Document;
        }

        #endregion Constructors

        #region Public binding properties

        /// <summary>
        /// Gets the associated word document.
        /// </summary>
        public Document WordDocument { get; private set; }

        /// <summary>
        /// Gets the list of all available test configurations
        /// </summary>
        public IList<ITfsTestConfiguration> TestConfigurations
        {
            get
            {

                if (SelectedTestPlan == null)
                    return _testConfigurations;
                _testConfigurations = new List<ITfsTestConfiguration> { new AllTestConfigurations() };
                //try
                //{

                foreach (var element in TestAdapter.GetAllTestConfigurationsForTestPlan(SelectedTestPlan))
                    _testConfigurations.Add(element);


                //foreach (var element in TestAdapter.GetAssignedTestConfigurationsForTestPlan(SelectedTestPlan, false,false))
                //    _testConfigurations.Add(element);
                //}
                //finally
                //{
                //    OnPropertyChanged("TestConfigurations");
                //    ViewDispatcher.BeginInvoke(new Action(() =>
                //    {
                //        var testConfigurationToSelect = _testConfigurations[0];
                //        if (!string.IsNullOrEmpty(StoredSelectedTestConfiguration))
                //        {
                //            if (_testConfigurations.Any(conf => conf.Name == StoredSelectedTestConfiguration))
                //                testConfigurationToSelect = _testConfigurations.First(conf => conf.Name == StoredSelectedTestConfiguration);
                //        }
                //        SelectedTestConfiguration = testConfigurationToSelect;
                //    }));
                //}


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
                {
                    return;
                }
                _selectedTestConfiguration = value;
                if (_selectedTestConfiguration != null)
                {
                    StoredSelectedTestConfiguration = _selectedTestConfiguration.Name;
                }
                OnPropertyChanged(nameof(SelectedTestConfiguration));
                if (ViewDispatcher != null)
                {
                    ViewDispatcher.Invoke(() => CreateReportCommand.CallEventCanExecuteChanged());
                }
            }
        }


        /// <summary>
        /// Gets or sets whether the test configurations are to include in generated document.
        /// </summary>
        public bool IncludeTestConfigurations
        {
            get
            {
                return _includeTestConfigurations;
            }
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
        /// Gets the list of all available test plans.
        /// </summary>
        public IList<ITfsTestPlan> TestPlans
        {
            get
            {
                if (_testPlans == null)
                {
                    StartBackgroundWorker(false, () =>
                    {
                        _testPlans = new List<ITfsTestPlan>(TestAdapter.AvailableTestPlans).OrderBy(o => o.Name).ToList();
                        OnPropertyChanged(nameof(TestPlans));
                        if (_testPlans != null && _testPlans.Count > 0 && ViewDispatcher != null)
                        {
                            ViewDispatcher.Invoke(() =>
                                {
                                    // Select one test plan
                                    var testPlanToSelect = _testPlans[0];
                                    if (!string.IsNullOrEmpty(StoredSelectedTestPlan) && _testPlans.Any(plan => plan.Name == StoredSelectedTestPlan))
                                    {

                                        testPlanToSelect = _testPlans.First(plan => plan.Name == StoredSelectedTestPlan);

                                    }
                                    SelectedTestPlan = testPlanToSelect;
                                });

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
                // Clear list for tree view
                _treeViewTestSuites.Clear();
                // Add the selected test plan to the collection for tree view if any selected
                if (_selectedTestPlan != null)
                {
                    _treeViewTestSuites.Add(_selectedTestPlan.RootTestSuite);
                }
                // No item in tree view selected
                SelectedTreeViewItem = null;
                if (CreateReportCommand != null)
                {
                    CreateReportCommand.CallEventCanExecuteChanged();
                    // Call begin invoke to select default test suite in tree view
                    if (ViewDispatcher != null)
                    {
                        ViewDispatcher.BeginInvoke(new Action(() =>
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
                                                                   SyncServiceTrace.I(string.Format("Exeption message: {0}. StackTrace: {1}.", e.Message, e.StackTrace));
                                                                   SyncServiceTrace.I(string.Format("Suite {0} cannot be found. Fallback to root suite of current plan {1}.", StoredSelectedTestSuite, SelectedTestPlan.Name));
                                                               }

                                                               SelectedTreeViewItem = testSuiteToSelect;
                                                               CreateReportCommand.CallEventCanExecuteChanged();
                                                           }
                                                       }));
                    }
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

                // on first value assignment use default value
                if (_selectedTestSuite == null && ViewDispatcher != null && ViewDispatcher.IsDispatching && _defaultTestSuiteName != null)
                {
                    var testSuiteSearcher = new TestSuiteSearcher(SelectedTestPlan);
                    var defaultTestSuite = testSuiteSearcher.SearchTestSuiteWithinTestPlan(_defaultTestSuiteName);
                    value = defaultTestSuite ?? value;
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
            get
            {
                return _skipLevels;
            }
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

        #endregion Public binding properties

        #region Private methods

        private void SetTestReportDefaults(ITestSpecReportDefault configuredDefaults)
        {
            if (configuredDefaults == null)
            {
                return;
            }

            //direclty initialized
            CreateDocumentStructure = configuredDefaults.CreateDocumentStructure;
            SelectedDocumentStructureType = configuredDefaults.DocumentStructureType;
            IncludeTestConfigurations = configuredDefaults.IncludeTestConfigurations;
            SelectedConfigurationPositionType = configuredDefaults.ConfigurationPositionType;
            SelectedTestCaseSortType = configuredDefaults.TestCaseSortType;
            SkipLevels = configuredDefaults.SkipLevels;

            //lazy initilized because value list is evalueated in background thread
            _defaultTestPlanName = configuredDefaults.SelectTestPlan;
            _defaultTestSuiteName = configuredDefaults.SelectTestSuite;
        }
        #endregion Private methods

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
            return !(SelectedTestPlan == null || SelectedTestSuite == null);
        }

        #endregion Private command methods

        #region Private create report methods

        /// <summary>
        /// The method creates report for selected test suite.
        /// </summary>
        public void CreateReport()
        {
            SyncServiceTrace.D("Creating report...");
            SyncServiceDocumentModel.TestReportRunning = true;
            try
            {
                ProcessPreOperations();

                // Common part
                if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.AboveTestPlan)
                {
                    CreateConfigurationPart();
                }
                if (CreateDocumentStructure)
                {
                    switch (SelectedDocumentStructureType)
                    {
                        case DocumentStructureType.AreaPath:
                            CreateReportByAreaPath(SelectedTestSuite);
                            break;
                        case DocumentStructureType.IterationPath:
                            CreateReportByIterationPath(SelectedTestSuite);
                            break;
                        case DocumentStructureType.TestPlanHierarchy:
                            CreateReportByTestPlanHierarchy(SelectedTestSuite);
                            break;
                    }
                }
                else CreateReportUnstructured(SelectedTestSuite);
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
                    SyncServiceTrace.I("Skipped creation of summary page and processing post operations because of error or cancellation.");
                }

            }
            catch (Exception ex)
            {
                var infoStorageService = SyncServiceFactory.GetService<IInfoStorageService>();
                if (infoStorageService == null) throw;
                IUserInformation info = new UserInformation
                {
                    Text = Resources.TestSpecification_Error,
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
            var templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestConfigurationElementTemplate;
            testReportHelper.InsertHeaderTemplate(config.ConfigurationTest.GetHeaderTemplate(templateName));

            // Determine configurations
            var configList = new List<ITfsTestConfiguration>();
            //if (SelectedTestConfiguration.Id == AllTestConfigurations.AllTestConfigurationsId)
            //{
            foreach (var testConfiguration in TestConfigurations)
            {
                if (testConfiguration.Id == AllTestConfigurations.AllTestConfigurationsId)
                {
                    continue;
                }
                configList.Add(testConfiguration);
            }
            //}
            //else
            //    configList.Add(SelectedTestConfiguration);
            // Iterate through configurations
            foreach (var testConfiguration in configList)
            {
                testReportHelper.InsertTestConfiguration(templateName, TestAdapter.GetTestConfigurationDetail(testConfiguration));
            }
        }


        /// <summary>
        /// The method creates unstructured report - writes out all test cases in one block.
        /// </summary>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> is the source of test cases to write out.</param>
        private void CreateReportUnstructured(ITfsTestSuite testSuite)
        {
            if (testSuite == null)
            {
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Get the detailed test plan.
            var testPlanDetail = TestAdapter.GetTestPlanDetail(testSuite);
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
            // Insert test plan template
            var templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestPlanTemplate;
            testReportHelper.InsertTestPlanTemplate(templateName, testPlanDetail);

            // Common part
            if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathTestPlan)
            {
                CreateConfigurationPart();
            }

            // Insert test suite template
            templateName = config.ConfigurationTest.ConfigurationTestSpecification.RootTestSuiteTemplate;
            var expandSharedSteps = config.ConfigurationTest.ExpandSharedSteps;
            testReportHelper.InsertTestSuiteTemplate(templateName, testSuiteDetail);

            // Get test cases of the test suite - with test cases in sub test suites
            var testCases = TestAdapter.GetAllTestCases(testSuiteDetail.TestSuite, expandSharedSteps);

            // TestCasesHelper need document structure, but the enumerations has not value 'None'
            // We will use the functionality without this structure capability
            var helper = new TestCaseHelper(testCases, SelectedDocumentStructureType);
            // Get sorted test cases
            var sortedTestCases = helper.GetTestCases(SelectedTestCaseSortType);
            // Test if test cases exists
            if (sortedTestCases != null && sortedTestCases.Count > 0)
            {
                // Write out the common part of test case block
                templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestCaseElementTemplate;
                testReportHelper.InsertHeaderTemplate(config.ConfigurationTest.GetHeaderTemplate(templateName));
                // Iterate all test cases
                foreach (var testCase in sortedTestCases)
                {
                    if (CancellationPending())
                    {
                        return;
                    }
                    // Write out the test case part
                    testReportHelper.InsertTestCase(templateName, testCase);
                }
            }
        }

        /// <summary>
        /// The method creates report structured by area path.
        /// </summary>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> is the source of test cases to write out.</param>
        private void CreateReportByAreaPath(ITfsTestSuite testSuite)
        {
            SyncServiceTrace.D("Creating report by area path...");
            if (testSuite == null)
            {
                SyncServiceTrace.D("Test suite is null");
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Get the detailed test plan.
            var testPlanDetail = TestAdapter.GetTestPlanDetail(testSuite);
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
            // Insert test plan template
            var templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestPlanTemplate;
            testReportHelper.InsertTestPlanTemplate(templateName, testPlanDetail);
            // Insert test suite template
            templateName = config.ConfigurationTest.ConfigurationTestSpecification.RootTestSuiteTemplate;
            testReportHelper.InsertTestSuiteTemplate(templateName, testSuiteDetail);

            //TODO MIS THIS IS THE ONLY PLACE WHERE GetAllTestCases can be called in context of the configuration template... Check late if it is enough only extend GetAllTestCases for at this single place here with the additional parameter
            // Get test cases of the test suite - with test cases in sub test suites
            var expandSharedSteps = config.ConfigurationTest.ExpandSharedSteps;
            var testCases = TestAdapter.GetAllTestCases(testSuiteDetail.TestSuite, expandSharedSteps);

            // Create test case helper
            var helper = new TestCaseHelper(testCases, SelectedDocumentStructureType);
            // Get all path elements from all test cases
            var pathElements = helper.GetPathElements(SkipLevels);

            // Iterate through path elements
            foreach (var pathElement in pathElements)
            {
                // Insert heading
                testReportHelper.InsertHeadingText(pathElement.PathPart, pathElement.Level - SkipLevels);
                // Get sorted test cases
                var sortedTestCases = helper.GetTestCases(pathElement, SelectedTestCaseSortType);
                SyncServiceTrace.D("Number of test cases:" + sortedTestCases.Count);
                if (sortedTestCases != null && sortedTestCases.Count > 0)
                {
                    // Write out the common part of test case block
                    templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestCaseElementTemplate;
                    testReportHelper.InsertHeaderTemplate(config.ConfigurationTest.GetHeaderTemplate(templateName));
                    // Iterate all test cases
                    foreach (var testCase in sortedTestCases)
                    {
                        if (CancellationPending())
                        {
                            return;
                        }
                        // Write out the test case part
                        testReportHelper.InsertTestCase(templateName, testCase);
                    }
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
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> is the source of test cases to write out.</param>
        private void CreateReportByIterationPath(ITfsTestSuite testSuite)
        {
            SyncServiceTrace.D("Creating report by iteration path...");
            if (testSuite == null)
            {
                SyncServiceTrace.D("Test suite is null");
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Get the detailed test plan.
            var testPlanDetail = TestAdapter.GetTestPlanDetail(testSuite);
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
            // Insert test plan template
            var templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestPlanTemplate;
            testReportHelper.InsertTestPlanTemplate(templateName, testPlanDetail);
            // Insert test suite template
            templateName = config.ConfigurationTest.ConfigurationTestSpecification.RootTestSuiteTemplate;
            testReportHelper.InsertTestSuiteTemplate(templateName, testSuiteDetail);

            // TODO MIS THIS IS THE ONLY PLACE WHERE GetAllTestCases can be called in context of the configuration template... Check late if it is enough only extend GetAllTestCases for at this single place here with the additional parameter
            // Get test cases of the test suite - with test cases in sub test suites
            var expandSharedSteps = config.ConfigurationTest.ExpandSharedSteps;
            var testCases = TestAdapter.GetAllTestCases(testSuiteDetail.TestSuite, expandSharedSteps);
            SyncServiceTrace.D("Number of test cases:" + testCases.Count);
            // Create test case helper
            var helper = new TestCaseHelper(testCases, SelectedDocumentStructureType);
            // Get all path elements from all test cases
            var pathElements = helper.GetPathElements(SkipLevels);

            // Iterate through path elements
            foreach (var pathElement in pathElements)
            {
                // Insert heading
                testReportHelper.InsertHeadingText(pathElement.PathPart, pathElement.Level - SkipLevels);
                // Get sorted test cases
                var sortedTestCases = helper.GetTestCases(pathElement, SelectedTestCaseSortType);
                if (sortedTestCases != null && sortedTestCases.Count > 0)
                {
                    // Write out the common part of test case block
                    templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestCaseElementTemplate;
                    testReportHelper.InsertHeaderTemplate(config.ConfigurationTest.GetHeaderTemplate(templateName));
                    // Iterate all test cases
                    foreach (var testCase in sortedTestCases)
                    {
                        if (CancellationPending())
                        {
                            return;
                        }
                        // Write out the test case part
                        testReportHelper.InsertTestCase(templateName, testCase);
                    }
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
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> is the source of test cases to write out.</param>
        private void CreateReportByTestPlanHierarchy(ITfsTestSuite testSuite)
        {
            SyncServiceTrace.D("Creating report by test plan hierarchy...");
            if (testSuite == null)
            {
                SyncServiceTrace.D("Test suite is null");
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Get the detailed test plan.
            var testPlanDetail = TestAdapter.GetTestPlanDetail(testSuite);
            if (testPlanDetail == null)
            {
                return;
            }

            // Create report helper
            var testReportHelper = new TestReportHelper(TestAdapter, testReport, config, CancellationPending);
            // Insert test plan template
            var templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestPlanTemplate;
            testReportHelper.InsertTestPlanTemplate(templateName, testPlanDetail);

            // Common part
            if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathTestPlan)
            {
                CreateConfigurationPart();
            }

            // Call recursion
            CreateReportByTestSuiteHierarchy(testSuite, 1, true);
        }

        /// <summary>
        /// The method generates report by test suite structure - recursion method called from CreateReportByTestPlanHierarchy.
        /// </summary>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> is the source of test cases to write out.</param>
        /// <param name="headingLevel">Level of the heading. First level is 1.</param>
        /// <param name="firstSuite"></param>
        private void CreateReportByTestSuiteHierarchy(ITfsTestSuite testSuite, int headingLevel, bool firstSuite)
        {
            SyncServiceTrace.D("Creating report by test suite hierarchy...");
            if (testSuite == null)
            {
                SyncServiceTrace.D("Test suite is null");
                return;
            }
            // Get the necessary objects.
            var config = SyncServiceFactory.GetService<IConfigurationService>().GetConfiguration(Document);
            var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(Document, config);
            // Get the detailed test suite.
            var testSuiteDetail = TestAdapter.GetTestSuiteDetail(testSuite);
            if (testSuiteDetail == null)
            {
                return;
            }

            var expandSharedSteps = config.ConfigurationTest.ExpandSharedSteps;
            // Check if this test suite or any child test suite contains at least one test case.
            // If no, to nothing.
            var allTestCases = TestAdapter.GetAllTestCases(testSuite, expandSharedSteps);
            if (allTestCases == null || allTestCases.Count == 0)
            {
                SyncServiceTrace.D("Given test suite does not contain tast cases.");
                return;
            }
            SyncServiceTrace.D("Number of test cases:" + allTestCases.Count);
            // Create report helper
            var testReportHelper = new TestReportHelper(TestAdapter, testReport, config, CancellationPending);
            // Insert heading
            testReportHelper.InsertHeadingText(testSuiteDetail.TestSuite.Title, headingLevel);
            var templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestSuiteTemplate;
            // Insert leaf test suite template for leaf suites
            if (testSuiteDetail.TestSuite.TestSuites == null || !testSuiteDetail.TestSuite.TestSuites.Any())
                templateName = config.ConfigurationTest.ConfigurationTestSpecification.LeafTestSuiteTemplate;
            // Insert root test suite template: if null, root test suite template = test suite template
            if (firstSuite)
            {
                templateName = config.ConfigurationTest.ConfigurationTestSpecification.RootTestSuiteTemplate;
            }
            testReportHelper.InsertTestSuiteTemplate(templateName, testSuiteDetail);

            // Print the test configuration beneath the first testsuite
            if (!_testConfigurationInformationPrinted && IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathFirstTestSuite)
            {
                CreateConfigurationPart();
                _testConfigurationInformationPrinted = true;
            }


            // Get test cases of the test suite - only test cases in this test suite
            var testCases = TestAdapter.GetTestCases(testSuiteDetail.TestSuite, expandSharedSteps);

            // TestCasesHelper need document structure, but the enumerations has not value 'None'
            // We will use the functionality without this structure capability
            var helper = new TestCaseHelper(testCases, SelectedDocumentStructureType);
            // Get sorted test cases
            var sortedTestCases = helper.GetTestCases(SelectedTestCaseSortType);
            // Test if test cases exists
            if (sortedTestCases != null && sortedTestCases.Count > 0)
            {
                // Write out the common part of test case block
                templateName = config.ConfigurationTest.ConfigurationTestSpecification.TestCaseElementTemplate;
                testReportHelper.InsertHeaderTemplate(config.ConfigurationTest.GetHeaderTemplate(templateName));
                // Iterate all test cases
                foreach (var testCase in sortedTestCases)
                {
                    if (CancellationPending())
                    {
                        return;
                    }
                    // Write out the test case part
                    testReportHelper.InsertTestCase(templateName, testCase);
                }
            }

            // Check child suites
            if (testSuiteDetail.TestSuite.TestSuites != null)
            {
                // Iterate through child test suites
                foreach (var childTestSuite in testSuiteDetail.TestSuite.TestSuites)
                {
                    if (CancellationPending())
                    {
                        return;
                    }
                    CreateReportByTestSuiteHierarchy(childTestSuite, headingLevel + 1, false);
                }
            }

            //Print the configuration information beneath all TestSuites
            if (IncludeTestConfigurations && SelectedConfigurationPositionType == Contracts.Enums.Model.ConfigurationPositionType.BeneathTestSuites)
            {
                CreateConfigurationPart();
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
            var operation = config.ConfigurationTest.GetPreOperationsForTestSpecification();
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
            var operation = config.ConfigurationTest.GetPostOperationsForTestSpecification();
            var testReportHelper = new TestReportHelper(TestAdapter, testReportOperation, config, CancellationPending);
            testReportHelper.ProcessOperations(operation);
            testReportOperation.UndoPreparationsDocumentForLongTermOperation();
        }
        #endregion Private create report methods
    }
}
