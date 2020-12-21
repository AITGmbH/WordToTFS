using System.Reflection;
using AIT.TFS.SyncService.Adapter.Word2007.Properties;
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
using AIT.TFS.SyncService.Contracts.Adapter;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;

namespace AIT.TFS.SyncService.Adapter.Word2007
{
    using Common.Helper;
    using Microsoft.TeamFoundation.ProcessConfiguration.Client;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;
    /// <summary>
    /// Base implementation of word 2007 sync adapter.
    /// </summary>
    internal abstract class Word2007SyncAdapter : IWordSyncAdapter
    {
        private readonly TypeOfInitialization _initializationType;
        private string _fileName;
        private Application _wordApplication;
        private PreparationDocumentHelper _preparationDocumentHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="Word2007SyncAdapter"/> class.
        /// </summary>
        /// <param name="configuration">Configuration to work with</param>
        /// <param name="document">Word document to work with.</param>
        public Word2007SyncAdapter(Document document, IConfiguration configuration)
        {
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            Configuration = configuration;
            _initializationType = TypeOfInitialization.ActiveDocument;
            Document = document;

            // These options need to be turned set. VML adds spaces (after removing the vml tags) that confuse the value comparision. Different image qualities too.
            document.WebOptions.RelyOnVML = false;
            document.ShowRevisions = false;
            document.WebOptions.AllowPNG = true;
            
            _preparationDocumentHelper = new PreparationDocumentHelper(Document, _wordApplication);

        }

        /// <summary>
        /// Gets the type of initialization of the class - either over file name or over active document.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        protected TypeOfInitialization InitializationType
        {
            get { return _initializationType; }
        }

        /// <summary>
        /// Gets or sets the name of file to work with. (See the <see cref="InitializationType"/>.)
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        protected string FileName
        {
            get { return _fileName; }
            set { _fileName = value; }
        }

        /// <summary>
        /// Gets the active document to work with. (See the <see cref="InitializationType"/>.)
        /// </summary>
        protected Document Document { get; private set; }

        /// <summary>
        /// Gets the configuration to work with.
        /// </summary>
        protected IConfiguration Configuration { get; private set; }

        #region IWorkItemSyncAdapter Members

        /// <summary>
        /// A collection of all opened work items. The collection will be initialized when calling open
        /// </summary>
        public IWorkItemCollection WorkItems { get; protected set; }

        /// <summary>
        /// Indicates whether this adapter has been opened. The list of work items will not be initialized unless
        /// an adapter is opened
        /// </summary>
        public bool IsOpen
        {
            get { return null != Document; }
        }

        /// <summary>
        /// The method executes all operations.
        /// </summary>
        /// <param name="operations">List of operations to execute.</param>
        public void ProcessOperations(IList<IConfigurationTestOperation> operations)
        {
            OperationsHelper.ProcessOperations(operations, Document);
        }

        /// <summary>
        /// Open the document for read and write of work items
        /// </summary>
        /// <param name="configurations">Ids of the work items to open and what configurations to use.</param>
        /// <returns>Returns true on success, false otherwise.</returns>
        public bool OpenWithConfigurations(Dictionary<int, IConfigurationItem> configurations )
        {
            if (configurations == null)
            {
                throw new ArgumentNullException("configurations");
            }

            return Open(configurations.Keys.ToArray());
        }

        /// <summary>
        /// Open the document for read and write of work items
        /// </summary>
        /// <param name="ids">Ids of the work items to open</param>
        /// <returns>Returns true on success, false otherwise</returns>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        public bool Open(int[] ids)
        {
            SyncServiceTrace.D(Resources.OpenWordAdapter);

            if (IsOpen && WorkItems != null)
            {
                return true;
            }

            if (_initializationType == TypeOfInitialization.ActiveDocument)
            {
                return LoadStructure(ids);
            }

            try
            {
                _wordApplication = new Application {Visible = true};
                if (null != _wordApplication)
                {
                    Document = FindOpenDocument();
                    if (null == Document)
                    {
                        Document = _wordApplication.Documents.Open(FileName: _fileName, ReadOnly: false, Visible: false);

                        _wordApplication.ActiveWindow.View.ShowRevisionsAndComments = false;
                        _wordApplication.ActiveWindow.View.RevisionsView = WdRevisionsView.wdRevisionsViewFinal;

                        return LoadStructure(ids);
                    }
                }
            }
            catch
            {
                Close();
                throw;
            }
            return false;
        }

        /// <summary>
        /// Not used for word adapter
        /// </summary>
        public bool Open(int[] ids, IList<string> fields)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not used for word adapter
        /// </summary>
        public bool Open(KeyValuePair<IConfigurationLinkItem,int[]> linkItem, IList<string> fields)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Closes the document and releases all resources
        /// </summary>
        public void Close()
        {
            if (!IsOpen || _initializationType != TypeOfInitialization.File)
            {
                return;
            }

            // Cast to _Document. In other case the Warning/Error is caused: Ambiguity between method '_Document.Close' and non-method 'DocumentEvent2_Event.Close'
            var document = Document as _Document;
            if (document != null)
            {
                document.Close(false);
            }

            Document = null;

            // Cast to _Application. In other case the Warning/Error is caused: Ambiguity between method '_Application.Quit' and non-method 'ApplicationEvents4_Event.Quit'
            var application = _wordApplication as _Application;
            if (application != null)
            {
                application.Quit(false);
            }

            _wordApplication = null;
            GC.Collect();
        }

        /// <summary>
        /// Creates a new <see cref="IWorkItem"/> object in this document.
        /// </summary>
        /// <param name="configuration">The configuration of the work item to create</param>
        /// <returns>The new <see cref="IWorkItem"/> object or null if the adapter failed to create work item.</returns>
        public IWorkItem CreateNewWorkItem(IConfigurationItem configuration)
        {
            Guard.ThrowOnArgumentNull(configuration, "configuration");

            SyncServiceTrace.D(Resources.CreateWorkItemBasedOnTemplate + configuration.RelatedTemplateFile);
            SyncServiceTrace.D("Assembly is " + Assembly.GetExecutingAssembly().Location);

            if (!File.Exists(configuration.RelatedTemplateFile))
            {
                SyncServiceTrace.E(Resources.LogService_Export_NoFile);
                return null;
            }

            // Check if cursor is not in table, add new paragraph
            var selection = Document.ActiveWindow.Selection;
            if (selection == null)
            {
                SyncServiceTrace.E(Resources.SelectionNotExists);
                return null;
            }

            if (WordSyncHelper.IsCursorBehindTable(selection))
                selection.TypeParagraph();

            if (WordSyncHelper.IsCursorInTable(selection))
            {
                return null;
            }

            // insert template. Selection is collapsed to end when inserting,
            // so note old start to get range where file is inserted
            var fullPath = Path.GetFullPath(configuration.RelatedTemplateFile);

            SyncServiceTrace.D(Resources.FullPathOfTemplateFile + fullPath);
            var oldStart = selection.Start;

            if (!File.Exists(fullPath))
            {
                SyncServiceTrace.E(Resources.LogService_Export_NoFile);
                return null;
            }


            selection.InsertFile(fullPath, ConfirmConversions: false);

            Range tableRange = Document.Range(oldStart, selection.End);

            Table table = tableRange.Tables.Cast<Table>().FirstOrDefault();
            if (table == null)
            {
                SyncServiceTrace.E(Resources.LogService_Export_NoTable);
            }
            else
            {
                // Add table marker if none is defined
                var tableRange2 = WordSyncHelper.GetCellRange(table, configuration.ReqTableCellRow, configuration.ReqTableCellCol);
                if (tableRange2 != null)
                {
                    if (tableRange2.Text.Replace("\r\a", string.Empty).Length == 0)
                    {
                        tableRange2.Text = configuration.WorkItemType + "\r\a";
                    }
                }

                var newItem = new WordTableWorkItem(table, configuration.WorkItemType, Configuration, configuration);
                WorkItems.Add(newItem);
                return newItem;
            }

            return null;
        }

        /// <summary>
        /// Saves all changes made to work items to its data source
        /// </summary>
        /// <returns>Returns a list of <see cref="ISaveError"/> if errors occurred. The list is empty if no error occurred</returns>
        public IList<ISaveError> Save()
        {
            if (IsOpen)
            {
                Document.Save();
            }

            return new List<ISaveError>();
        }

        /// <summary>
        /// Saves all new work items
        /// </summary>
        /// <returns>
        /// Returns a list of <see cref="ISaveError" /> if errors occurred. The list is empty if no error occurred
        /// </returns>
        public IList<ISaveError> SaveNew()
        {
            if (IsOpen)
            {
                Document.Save();
            }

            return new List<ISaveError>();
        }

        /// <summary>
        /// Validates all work items of the adapter to meet designated rules
        /// </summary>
        /// <returns>Returns a list of <see cref="ISaveError"/> if errors occurred. The list is empty if no error occurred</returns>
        [SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "links")]
        public IList<ISaveError> ValidateWorkItems()
        {
            var errors = new List<ISaveError>();

            var doubles = WorkItems.Where(x => WorkItems.Any(y => y.Id == x.Id && y.Id > 0 && y != x)).Distinct();
            foreach (var item in doubles)
            {
                errors.Add(new SaveError
                               {
                                   Exception = new InvalidOperationException(Resources.ExceptionText_DoubleIdsExplanation),
                                   Field = null,
                                   WorkItem = item
                               });
            }

            // check if all links are valid numbers
            foreach (var workItem in WorkItems)
            {
                try
                {
                    // ReSharper disable UnusedVariable
                    var links = workItem.Links;
                    // ReSharper restore UnusedVariable
                }
                catch (ConfigurationException ce)
                {
                    errors.Add(new SaveError { Exception = ce, Field = null, WorkItem = workItem });
                }
            }

            ValidateTableStructures(errors);

            return errors;
        }

        /// <summary>
        /// Check if tables are exactly as defined in the template files and not accidentally merged
        /// </summary>
        private void ValidateTableStructures(List<ISaveError> errors)
        {
            var tableStructureCache = new Dictionary<string, Dictionary<int, int>>();
            var application = (_Application) new Application();
            foreach (var workItem in WorkItems)
            {
                var wordTableWorkItem = workItem as WordTableWorkItem;
                if (wordTableWorkItem == null)
                {
                    continue;
                }

                // cache table structure so we dont have to open them for every work item
                if (tableStructureCache.ContainsKey(wordTableWorkItem.Configuration.RelatedTemplateFile) == false)
                {
                    var structure = new Dictionary<int, int>();
                    var document = application.Documents.Open(Path.GetFullPath(wordTableWorkItem.Configuration.RelatedTemplateFile), AddToRecentFiles: false, Visible: false);
                    for (int rowIndex = 1; rowIndex <= document.Tables[1].Rows.Count; rowIndex++)
                    {
                        structure.Add(rowIndex, document.Tables[1].Rows[rowIndex].Cells.Count);
                    }

                    document.Close(SaveChanges: false);
                    tableStructureCache.Add(wordTableWorkItem.Configuration.RelatedTemplateFile, structure);
                }

                var cachedStructure = tableStructureCache[wordTableWorkItem.Configuration.RelatedTemplateFile];
                var rows = wordTableWorkItem.Table.Rows;

                if (cachedStructure.Keys.Count != rows.Count)
                {
                    errors.Add(new SaveError
                                   {
                                       Exception = new InvalidOperationException(Resources.ExceptionText_SkipInvalidTable),
                                       Field = null,
                                       WorkItem = workItem
                                   });
                    continue;
                }

                for (int rowIndex = 1; rowIndex <= cachedStructure.Keys.Count; rowIndex++)
                {
                    if (cachedStructure[rowIndex] != rows[rowIndex].Cells.Count)
                    {
                        errors.Add(new SaveError
                                       {
                                           Exception = new InvalidOperationException(Resources.ExceptionText_SkipInvalidTable),
                                           Field = null,
                                           WorkItem = workItem
                                       });
                        break;
                    }
                }
            }

            application.Quit(SaveChanges: false);
        }

        /// <summary>
        /// No used. Word document support no web access
        /// </summary>
        public Uri GetWorkItemEditorUrl(int id)
        {
            return null;
        }

        #endregion

        /// <summary>
        /// Search the document in all opened documents.
        /// </summary>
        /// <returns>Found document, otherwise null.</returns>
        private Document FindOpenDocument()
        {
            foreach (Document document in _wordApplication.Documents)
            {
                if (_fileName.Equals(document.Path))
                {
                    return document;
                }
            }

            return null;
        }

        /// <summary>
        /// Method defined as abstract to perform the specific load operation in derived classes.
        /// </summary>
        /// <param name="ids">The ids of the work items to load</param>
        /// <returns>Result of the operation.</returns>
        protected abstract bool LoadStructure(int[]ids);

        /// <summary>
        /// Defines the manner of creation of the class <see cref="Word2007SyncAdapter"/> and derived classes.
        /// </summary>
        protected enum TypeOfInitialization
        {
            File,
            ActiveDocument
        }

        /// <summary>
        /// Gets the work items represented by the structures in the current selection
        /// </summary>
        /// <returns>
        /// The currently selected work items
        /// </returns>
        public abstract IEnumerable<IWorkItem> GetSelectedWorkItems();

        /// <summary>
        /// Gets the header that is currently selected
        /// </summary>
        /// <returns></returns>
        public abstract IEnumerable<IWorkItem> GetSelectedHeader();

        /// <summary>
        /// Creates a work item hyperlink at the current position in the document.
        /// </summary>
        /// <param name="text">Text for the hyperlink</param>
        /// <param name="uri">URI to the team server web access.</param>
        public void CreateWorkItemHyperlink(string text, Uri uri)
        {
            Guard.ThrowOnArgumentNull(text, "text");
            Guard.ThrowOnArgumentNull(uri, "uri");

            var wrdSelection = Document.ActiveWindow.Selection;
            wrdSelection.Hyperlinks.Add(Anchor: wrdSelection.Range, Address: uri.AbsoluteUri, TextToDisplay: text);
            wrdSelection.InsertAfter("\r\n");
            Document.ActiveWindow.Selection.SetRange(wrdSelection.End +1, wrdSelection.End + 1);
        }

        /// <summary>
        /// Prepare the document for the long-term operation. 
        /// Disable Spellchecking and Pagination. This is necessary, because Word Add-ins are not supposed to create big documents in a self running manner
        /// </summary>
        public void PrepareDocumentForLongTermOperation()
        {
            _preparationDocumentHelper.PrepareDocumentForLongTermOperation();
        }

        /// <summary>
        /// Undo the changes to the document 
        /// Enable Spellchecking and Pagination.
        /// </summary>
        public void UndoPreparationsDocumentForLongTermOperation()
        {
            _preparationDocumentHelper.UndoPreparationsDocumentForLongTermOperation();
        }

        /// <summary>
        /// Gets the list of available work items from tfs.
        /// </summary>
        public ICollection<WorkItem> AvailableWorkItemsFromTFS
        {
            get
            {
                return null;
            }
        }


        /// <summary>
        /// A collection of work item colors.
        /// </summary>
        public IDictionary<string, WorkItemTypeAppearance> WorkItemColors
        {
            get
            {
                return null;
            }
        }
    }
}
