#region Usings
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Threading;
using AIT.TFS.SyncService.Adapter.TFS2012;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Adapter;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Model.Properties;
using AIT.TFS.SyncService.Model.TestReport;
using AIT.TFS.SyncService.Model.WindowModel;
using AIT.TFS.SyncService.Service.Configuration;
using AIT.TFS.SyncService.Service.Configuration.Serialization.Console;
using Microsoft.Office.Interop.Word;
using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
#endregion

namespace AIT.TFS.SyncService.Model.Console
{
    /// <summary>
    /// Console Helper that handles the document creation with the console extension
    /// </summary>
    public class ConsoleExtensionHelper
    {
        #region Fields

        // All elements that are necessary for the creation of a document including the document model to the save the work items
        private Document _wordDocument;
        private ISyncServiceDocumentModel _documentModel;
        private Application _wordApplication;

        // Internal helpers
        private DocumentConfiguration _documentConfiguration;
        private List<int> _workItems;
        private bool _testSpecByQuery;

        private IWorkItemSyncService _workItemSyncService;
        private IConfigurationService _configService;
        private IConfiguration _configuration;
        private IWorkItemSyncAdapter _source;
        private IWordSyncAdapter _destination;

        private ITfsService _tfsService;

        // The test helpers
        private string _fileName;
        // ReSharper disable once InconsistentNaming
        private const string TIME_STAMP_VARIABLE = "%TIMESTAMP%";

        // other
        private readonly ITestReportingProgressCancellationService _testReportingProgressCancellationService;
        private CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();
        public const char ReportingSuitePathDelimiter = '/';

        #endregion

        #region Contstructors

        public ConsoleExtensionHelper(ITestReportingProgressCancellationService testReportingProgressCancellationService)
        {
            _testReportingProgressCancellationService = testReportingProgressCancellationService;
        }
        #endregion Contstructors

        #region Properties

        /// <summary>
        /// Determines if the helper is importing at the moment
        /// </summary>
        public bool IsImporting
        {
            get;
            private set;
        }

        /// <summary>
        /// IsFindingWorkItems
        /// </summary>
        private bool IsFindingWorkItems { get; set; }

        /// <summary>
        /// Gets collection of all found WorkItems.
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Justification = "You are not supposed to manipulate the list.")]
        private ICollection<DataItemModel<SynchronizedWorkItemViewModel>> FoundWorkItems
        {
            get;
            set;
        }

        #endregion

        #region Public methods

        /// <summary>
        ///  This method creates a word document with the GetWorkItems-Command
        /// </summary>
        /// <param name="documentConfiguration">The configuration</param>
        /// <param name="workItems">The string that identifies the work items that should be imported</param>
        /// <param name="type">The type of the import operation</param>
        public void CreateWorkItemDocument(DocumentConfiguration documentConfiguration, string workItems, string type)
        {
            try
            {
                _documentConfiguration = documentConfiguration;

                // Initialize the document
                Initialize();

                if (type.Equals("ByID"))
                {
                    _workItems = GetWorkItemsIdsByString(workItems);
                }
                else if (type.Equals("ByQuery"))
                {
                    _workItems = GetWorkItemsIdsByQueryName(workItems);
                }

                if (_workItems == null || _workItems.Count < 1)
                {
                    SyncServiceTrace.E(Resources.NoWorkItemsFoundInformation);
                    throw new Exception(Resources.NoWorkItemsFoundInformation);
                }

                // Get the work items that should be imported / This will come from the query settings
                Find();

                while (IsFindingWorkItems)
                {
                    Thread.CurrentThread.Join(50);
                }

                // Show how many work items were found
                ConsoleExtensionLogging.LogMessage(string.Format("{0}{1}", Resources.WorkItemsCountInformation, FoundWorkItems.Count), ConsoleExtensionLogging.LogLevel.Both);
                SyncServiceTrace.D(string.Format("{0}{1}", Resources.WorkItemsCountInformation, FoundWorkItems.Count));

                Import();

                while (IsImporting)
                {
                    Thread.CurrentThread.Join(50);
                }

                ConsoleExtensionLogging.LogMessage(Resources.WorkItemsImportedInformation, ConsoleExtensionLogging.LogLevel.Both);
                SyncServiceTrace.D(Resources.WorkItemsImportedInformation);

                FinalizeDocument();
            }
            catch (Exception e)
            {
                // close any open word applications form this method
                if (_wordApplication != null)
                {
                    _wordApplication.Quit();
                }
                SyncServiceTrace.LogException(e);
                // ReSharper disable once PossibleIntendedRethrow
                throw e;
            }
        }

        /// <summary>
        /// This method creates a word document with the TestSpecificationReport-Command
        /// </summary>
        public void CreateTestSpecDocument(DocumentConfiguration documentConfiguration)
        {
            try
            {
                SyncServiceTrace.D(Resources.CreateTestSpecificationDoc);
                _documentConfiguration = documentConfiguration;
                Initialize();
                var testSpecSettings = _documentConfiguration.TestSpecSettings;
                var testAdapter = SyncServiceFactory.CreateTfsTestAdapter(_documentModel.TfsServer, _documentModel.TfsProject, _documentModel.Configuration);
                TestCaseSortType testCaseSortType;
                Enum.TryParse(testSpecSettings.SortTestCasesBy, out testCaseSortType);
                DocumentStructureType documentStructure;
                Enum.TryParse(testSpecSettings.DocumentStructure, out documentStructure);
                ITfsTestPlan testPlan = testAdapter.GetTestPlans(null).Where(plan => plan.Name == testSpecSettings.TestPlan).FirstOrDefault();
                if (testPlan == null)
                {
                    SyncServiceTrace.D(Resources.TestSettingsInvalidTestPlan);
                    throw new ArgumentException(Resources.TestSettingsInvalidTestPlan);
                }

                var testSuiteSearcher = new TestSuiteSearcher(testPlan);
                var testSuite = testSuiteSearcher.SearchTestSuiteWithinTestPlan(testSpecSettings.TestSuite);

                var reportModel = new TestSpecificationReportModel(_documentModel,
                    testAdapter,
                    testPlan,
                    testSuite,
                    testSpecSettings.CreateDocumentStructure,
                    testSpecSettings.SkipLevels,
                    documentStructure,
                    testCaseSortType,
                    new TestReportingProgressCancellationService(false));
                reportModel.CreateReport();

                FinalizeDocument();
            }
            catch (Exception e)
            {
                if (_wordApplication != null)
                {
                    _wordApplication.Quit();
                }
                // ReSharper disable once PossibleIntendedRethrow
                throw e;
            }
        }

        /// <summary>
        /// This method creates a word document with the TestSpecReportByQuery-Command
        /// </summary>
        /// <param name="documentConfiguration"></param> Document-Configuration
        /// <param name="workItems"></param> In case of type=ByQuery querypath, else workitem-Ids as string
        /// <param name="type"></param> {ByQuery, ById}
        public void CreateTestSpecDocumentByQuery(DocumentConfiguration documentConfiguration, string workItems, string type)
        {
            _testSpecByQuery = true;
            try
            {
                CreateWorkItemDocument(documentConfiguration, workItems, type);
            }
            catch (Exception e)
            {
                if (_wordApplication != null)
                {
                    _wordApplication.Quit();
                }
                // ReSharper disable once PossibleIntendedRethrow
                throw e;
            }
        }

        public void AutoRefreshGivenQuery(string queryName, Document wordDocument, ISyncServiceDocumentModel documentModel, IConfiguration configuration)
        {

            //RegisterServices();

            _documentModel = documentModel;
            _wordDocument = documentModel.WordDocument as Document;

            _configuration = configuration;
            _documentModel.Configuration.RefreshMappings();
            _documentModel.Configuration.ActivateMapping(_documentModel.MappingShowName);

            //Get the service
            GetService();

            //Search the workitems
            _workItems = GetWorkItemsIdsByQueryName(queryName);

            Find();

            while (IsFindingWorkItems)
            {
                Thread.CurrentThread.Join(50);
            }

            // Show how many work items were found
            ConsoleExtensionLogging.LogMessage(string.Format("{0}{1}", Resources.WorkItemsCountInformation, FoundWorkItems.Count), ConsoleExtensionLogging.LogLevel.Both);
            SyncServiceTrace.D((string.Format("{0}{1}", Resources.WorkItemsCountInformation, FoundWorkItems.Count)));

            Import();
            while (IsImporting)
            {
                Thread.CurrentThread.Join(50);
            }
        }

        /// <summary>
        /// This method creates a word document with the TestResultReport-Command
        /// </summary>
        public void CreateTestResultDocument(DocumentConfiguration documentConfiguration)
        {
            try
            {
                SyncServiceTrace.D(Resources.CreateTestResultDoc);
                _documentConfiguration = documentConfiguration;
                Initialize();
                var testResultSettings = _documentConfiguration.TestResultSettings;
                var testAdapter = SyncServiceFactory.CreateTfsTestAdapter(_documentModel.TfsServer, _documentModel.TfsProject, _documentModel.Configuration);
                TestCaseSortType testCaseSortType;
                Enum.TryParse(testResultSettings.SortTestCasesBy, out testCaseSortType);
                DocumentStructureType documentStructure;
                Enum.TryParse(testResultSettings.DocumentStructure, out documentStructure);
                ConfigurationPositionType configurationPositionType;
                Enum.TryParse(testResultSettings.TestConfigurationsPosition, out configurationPositionType);
                var testPlan = testAdapter.GetTestPlans(null).Where(plan => plan.Name == testResultSettings.TestPlan).FirstOrDefault();
                if (testPlan == null)
                {
                    SyncServiceTrace.D(Resources.TestSettingsInvalidTestPlan);
                    throw new ArgumentException(Resources.TestSettingsInvalidTestPlan);
                }
                var testSuiteSearcher = new TestSuiteSearcher(testPlan);
                var testSuite = testSuiteSearcher.SearchTestSuiteWithinTestPlan(testResultSettings.TestSuite);

                var reportModel = new TestResultReportModel(_documentModel,
                    testAdapter,
                    testPlan,
                    testSuite,
                    testResultSettings.CreateDocumentStructure,
                    testResultSettings.SkipLevels,
                    documentStructure,
                    testCaseSortType,
                    testResultSettings.IncludeTestConfigurations,
                    configurationPositionType,
                    testResultSettings.IncludeOnlyMostRecentResults,
                    testResultSettings.MostRecentForAllConfigurations,
                    testResultSettings.Build,
                    testResultSettings.TestConfiguration,
                    new TestReportingProgressCancellationService(false));

                reportModel.CreateReport();
                FinalizeDocument();
            }
            catch
            {
                if (_wordApplication != null)
                {
                    _wordApplication.Quit();
                }
                // ReSharper disable once PossibleIntendedRethrow
                throw;
            }
        }



        #endregion

        #region Private methods

        /// <summary>
        /// Convert a string to a list of work Items
        /// </summary>
        /// <param name="workitems"></param>
        /// <returns>ids of workitems</returns>
        private List<int> GetWorkItemsIdsByString(string workitems)
        {
            ConsoleExtensionLogging.LogMessage(Resources.GetWorkItemById, ConsoleExtensionLogging.LogLevel.Both);
            SyncServiceTrace.D(Resources.GetWorkItemById);

            var stringIDs = workitems.Split(new[] { ' ', ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            int intValue;
            return stringIDs.Where(x => int.TryParse(x, out intValue)).Select(int.Parse).ToList();
        }

        /// <summary>
        /// Get all work item ids for a given query name
        /// </summary>
        /// <param name="query"></param>
        /// <returns>ids of workitems</returns>
        private List<int> GetWorkItemsIdsByQueryName(string query)
        {

            ConsoleExtensionLogging.LogMessage(Resources.GetWorkItemsByQueryName + query, ConsoleExtensionLogging.LogLevel.Both);
            SyncServiceTrace.D(Resources.GetWorkItemsByQueryName + query);
            var queryconfig = new QueryConfiguration
            {
                QueryPath = query,
                ImportOption = QueryImportOption.SavedQuery
            };
            var adapter = new Tfs2012SyncAdapter(_documentModel.TfsServer, _documentModel.TfsProject, queryconfig, _documentModel.Configuration);
            var intIDs = new List<int>();
            var queryWorkItems = adapter.LoadWorkItemsFromSavedQuery(queryconfig);
            foreach (var queryWorkItem in queryWorkItems)
            {
                intIDs.Add(queryWorkItem.Id);
            }

            return intIDs;
        }

        /// <summary>
        /// Finalize the document
        /// </summary>
        private void FinalizeDocument()
        {
            SyncServiceTrace.D(Resources.FinalizationOfDoc);
            UndoPreparationsDocumentForReportGeneration();
            _wordDocument.SaveAs(_fileName);

            ConsoleExtensionLogging.LogMessage(string.Format(Resources.FileSavedInformation, _fileName), ConsoleExtensionLogging.LogLevel.Both);
            SyncServiceTrace.D(string.Format(Resources.FileSavedInformation, _fileName));
            if (_documentConfiguration.Settings.CloseOnFinish)
            {
                _wordDocument.Close(true);
                _wordApplication.Quit();
            }

            ConsoleExtensionLogging.LogMessage(Resources.WordClosed, ConsoleExtensionLogging.LogLevel.Both);
            SyncServiceTrace.D(Resources.WordClosed);
        }


        /// <summary>
        /// Save the document
        /// </summary>
        private void SaveDocument()
        {
            if (_fileName == null)
            {
                InitializeFileNameWithTimeStamp();
            }

            // Filename cannot be null
            // ReSharper disable AssignNullToNotNullAttribute
            if (!Directory.Exists(Path.GetDirectoryName(_fileName)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_fileName));
            }
            // ReSharper restore AssignNullToNotNullAttribute

            if (_documentConfiguration.Settings.Overwrite)
            {
                // Save the document
                _wordDocument.SaveAs(_fileName);
            }
            else
            {
                if (File.Exists(_fileName))
                {
                    throw new Exception(Resources.FileExistsError);
                }

                // Save the document
                _wordDocument.SaveAs(_fileName);
            }
            SyncServiceTrace.D(Resources.WordDocSaveMessage);

        }


        /// <summary>
        /// Initialize the helper to create any document.
        /// </summary>
        private void Initialize()
        {
            try
            {
                SyncServiceTrace.D(Resources.InitializationOfDocument);

                InitializeFileNameWithTimeStamp();

                RegisterServices();

                // Create the Model
                CreateModel();

                // Configure;
                Configure();

                // Get the service
                GetService();
            }
            catch (Exception e)
            {
                // ReSharper disable once PossibleIntendedRethrow
                SyncServiceTrace.LogException(e);
                throw e;
            }

        }


        /// <summary>
        /// Initialize the filename for the current document.
        /// If the settings contain a variable for the timestamp, the the filename or folder will contain the current date and time.
        /// </summary>
        private void InitializeFileNameWithTimeStamp()
        {
            SyncServiceTrace.D(Resources.InitializationOfFileName);
            _fileName = _documentConfiguration.Settings.Filename;

            if (_documentConfiguration.Settings.Filename.Contains(TIME_STAMP_VARIABLE))
            {
                var currentTimeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                _fileName = _documentConfiguration.Settings.Filename.Replace(TIME_STAMP_VARIABLE, currentTimeStamp);
            }
        }

        /// <summary>
        /// Configure the Document model and load the specific template
        /// </summary>
        private void Configure()
        {
            _documentModel.Configuration.RefreshMappings();
            // Activate Mapping
            _documentModel.Configuration.ActivateMapping(_documentConfiguration.Settings.Template);
        }

        /// <summary>
        /// Create the model of the document and bind the project to the server
        /// </summary>
        private void CreateModel()
        {
            SyncServiceTrace.D(Resources.CreateModelInfo);
            // Create a word appliaction and set it to visible
            if (!_documentConfiguration.Settings.Overwrite && File.Exists(_fileName))
            {
                throw new Exception(Resources.FileExistsError);
            }
            SyncServiceTrace.D(Resources.FileIsOverwrittenOrNotExists);

            string wordTemplateFile = null;
            if (!string.IsNullOrEmpty(_documentConfiguration.Settings.DotxTemplate))
            {
                wordTemplateFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _documentConfiguration.Settings.DotxTemplate);

                // if the path is absolute, use it
                if (Path.IsPathRooted(_documentConfiguration.Settings.DotxTemplate))
                {
                    wordTemplateFile = _documentConfiguration.Settings.DotxTemplate;
                }
            }

            if (!string.IsNullOrEmpty(wordTemplateFile) && !File.Exists(wordTemplateFile))
            {
                throw new Exception(string.Format(Resources.Error_TemplateDoesNotExist, wordTemplateFile));
            }

            _wordApplication = new Application();
            _wordApplication.Visible = !_documentConfiguration.Settings.WordHidden;

            if (!string.IsNullOrEmpty(wordTemplateFile))
            {
                _wordApplication.Documents.Add(wordTemplateFile);
            }
            else
            {
                _wordApplication.Documents.Add();
            }


            _wordDocument = _wordApplication.ActiveDocument;
            PrepareDocumentForReportGeneration();
            _documentModel = new SyncServiceDocumentModel(_wordDocument);

            _documentModel.MappingShowName = _documentConfiguration.Settings.Template;
            _documentModel.BindProject(_documentConfiguration.Settings.Server, _documentConfiguration.Settings.Project);
            SaveDocument();
        }


        private static void RegisterServices()
        {
            SyncServiceTrace.D(Resources.RegisterServices);
            // Initialize all assemblies.
            AssemblyInit.Instance.Init();
            // ReSharper disable RedundantNameQualifier
            Adapter.Word2007.AssemblyInit.Instance.Init();
            Service.AssemblyInit.Instance.Init();
            AssemblyInit.Instance.Init();
            // ReSharper restore RedundantNameQualifier
        }

        /// <summary>
        /// Get the necessary services
        /// </summary>
        private void GetService()
        {
            _workItemSyncService = SyncServiceFactory.GetService<IWorkItemSyncService>();
            _configService = SyncServiceFactory.GetService<IConfigurationService>();
            _configuration = _configService.GetConfiguration(_wordDocument);


            _tfsService = SyncServiceFactory.CreateTfsService(_documentModel.TfsServer, _documentModel.TfsProject, _documentModel.Configuration);
            _source = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(_documentModel.TfsServer, _documentModel.TfsProject, null, _configuration);
            _destination = SyncServiceFactory.CreateWord2007TableWorkItemSyncAdapter(_wordDocument, _configuration);

        }

        /// <summary/>
        ///// Import the workitems to the document
        ///// </summary>
        private void Import()
        {
            // start background thread to execute the import
            var backgroundWorker = new BackgroundWorker();

            IsImporting = true;
            backgroundWorker.DoWork += DoImport;
            backgroundWorker.RunWorkerCompleted += ImportFinished;
            backgroundWorker.RunWorkerAsync(backgroundWorker);
        }

        /// <summary/>
        ///// Import the workitems to the document
        ///// </summary>
        private void Find()
        {
            _destination.Open(null);
            // start background thread to execute the import
            var backgroundWorker = new BackgroundWorker();
            IsFindingWorkItems = true;
            backgroundWorker.DoWork += DoFind;
            backgroundWorker.RunWorkerCompleted += FindingFinished;
            backgroundWorker.RunWorkerAsync(backgroundWorker);
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
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.HideProgress();

            SyncServiceTrace.I(Resources.WorkItemsAreImported);
        }

        /// <summary>
        /// Occurs when the background operation has completed, has been canceled, or has raised an exception.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A <see cref="RunWorkerCompletedEventArgs"/> that contains the event data.</param>
        private void FindingFinished(object sender, RunWorkerCompletedEventArgs e)
        {
            e.Error.NotifyIfException("Failed to find work items.");

            IsFindingWorkItems = false;

            SyncServiceTrace.I(Resources.WorkItemsAreFound);
        }

        /// <summary>
        /// Background method to import the work items from TFS.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A <see cref="RunWorkerCompletedEventArgs"/> that contains the event data</param>
        private void DoImport(object sender, DoWorkEventArgs e)
        {
            SyncServiceTrace.I(Resources.ImportWorkItems);

            if (_workItemSyncService == null || _configService == null)
            {
                SyncServiceTrace.D(Resources.ServicesNotExists);
                throw new Exception(Resources.ServicesNotExists);
            }

            _documentModel.Save();

            _destination.ProcessOperations(_configuration.PreOperations);
            var importWorkItems = (from wim in FoundWorkItems

                                   select wim.Item).ToList();
            SyncServiceTrace.D(Resources.NumberOfFoundWorkItems + FoundWorkItems.Count);

            // display progress dialog
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.ShowProgress();
            progressService.NewProgress(Resources.GetWorkItems_Title);

            // search for already existing work items in word and ask whether to overide them
            if (!(_source.Open(importWorkItems.Select(x => x.TfsItem.Id).ToArray()) && _destination.Open(null)))
            {
                return;
            }
            if (importWorkItems.Select(x => x.TfsItem.Id).Intersect(_destination.WorkItems.Select(x => x.Id)).Any()) { }
            if (!_testSpecByQuery)
            {
                _workItemSyncService.Refresh(_source, _destination, importWorkItems.Select(x => x.TfsItem), _configuration);
            }

            if (_testSpecByQuery)
            {
                var testAdapter = SyncServiceFactory.CreateTfsTestAdapter(_documentModel.TfsServer, _documentModel.TfsProject, _documentModel.Configuration);
                var model = new TestSpecificationReportByQueryModel(_tfsService, _documentModel, testAdapter, _workItems);
                var expandSharedsteps = _configuration.ConfigurationTest.ExpandSharedSteps;
                var testCases = model.GetTestCasesFromWorkItems(importWorkItems, expandSharedsteps);
                var sharedSteps = model.GetSharedStepsFromWorkItems(importWorkItems);
                // Get any pre and post operations for the reports
                var preOperations = _configuration.ConfigurationTest.GetPreOperationsForTestSpecification();
                var postOperations = _configuration.ConfigurationTest.GetPostOperationsForTestSpecification();
                var testReport = SyncServiceFactory.CreateWord2007TestReportAdapter(_wordDocument, _configuration);
                var testReportHelper = new TestReportHelper(testAdapter, testReport, _configuration, () => !IsImporting);

                testReportHelper.ProcessOperations(preOperations);
                if (sharedSteps.Count > 0)
                {
                    _workItemSyncService.RefreshAndSubstituteSharedStepItems(_source, _destination, importWorkItems.Select(x => x.TfsItem), _configuration, testReportHelper, sharedSteps);
                }
                if (testCases.Count > 0)
                {
                    _workItemSyncService.RefreshAndSubstituteTestItems(_source, _destination, importWorkItems.Select(x => x.TfsItem), _configuration, testReportHelper, testCases);
                }
                if (_testReportingProgressCancellationService.CheckIfContinue())
                {
                    testReportHelper.CreateSummaryPage(_wordDocument, null);
                    testReportHelper.ProcessOperations(postOperations);
                }
                else
                {
                    SyncServiceTrace.I(Resources.DoImportError);
                }

                _documentModel.TestReportGenerated = true;
            }

            _destination.ProcessOperations(_configuration.PostOperations);
        }


        /// <summary>
        /// Background method to find the work items from TFS.
        /// </summary>
        private void DoFind(object sender, DoWorkEventArgs e)
        {
            SyncServiceTrace.D(Resources.FindWorkItems);

            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.ShowProgress();
            progressService.NewProgress(Resources.GetWorkItems_Title);

            var queryConfiguration  = new QueryConfiguration();

            foreach (var workItem in _workItems)
            {
                queryConfiguration.ByIDs.Add(workItem);
            }

            queryConfiguration.ImportOption = QueryImportOption.IDs;

            _source.Open(queryConfiguration.ByIDs.ToArray());

            EnsureTfsLazyLoadingFinished(_source);

            _cancellationTokenSource.Cancel();
            _cancellationTokenSource = new CancellationTokenSource();
            var token = _cancellationTokenSource.Token;

            FoundWorkItems = (from wi in _source.WorkItems
                              select new DataItemModel<SynchronizedWorkItemViewModel>(new SynchronizedWorkItemViewModel(wi, null, _destination, _documentModel, token))
                              {
                                  IsChecked = true
                              }).ToList();
        }


        /// <summary>
        /// This beautiful piece of code ensures, that the TFS API has finished creating its field collection.
        /// Not doing this sequentially before creating parallel synchronization tasks will result in a
        /// "Collection has been modified" exception.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "x")]
        private static void EnsureTfsLazyLoadingFinished(IWorkItemSyncAdapter tfsService)
        {
            if (tfsService.WorkItems.Any())
            {
                foreach (var field in tfsService.WorkItems.First().Fields)
                {
                    // ReSharper disable once UnusedVariable
                    var x = field.Value;
                }
            }
        }


        /// <summary>
        /// Prepare the document for the generation of a report.
        /// Disable Spellchecking and Pagination. This is necessary, because Word Add-ins are not supposed to create big documents in a self running manner
        /// </summary>
        public void PrepareDocumentForReportGeneration()
        {
            _wordDocument.Application.Options.Pagination = false;
            _wordDocument.ShowGrammaticalErrors = false;
            _wordDocument.ShowSpellingErrors = false;
            _wordApplication.Application.Options.Pagination = false;
            _wordApplication.Application.Options.CheckGrammarAsYouType = false;
            _wordApplication.Application.Options.CheckSpellingAsYouType = false;
        }

        /// <summary>
        /// Undo the changes to the document
        /// Enable Spellchecking and Pagination.
        /// </summary>
        public void UndoPreparationsDocumentForReportGeneration()
        {
            _wordDocument.Application.Options.Pagination = true;
            _wordDocument.ShowGrammaticalErrors = true;
            _wordDocument.ShowSpellingErrors = true;
            _wordApplication.Application.Options.Pagination = true;
            _wordApplication.Application.Options.CheckGrammarAsYouType = true;
            _wordApplication.Application.Options.CheckSpellingAsYouType = true;
        }

        #endregion
    }
}
