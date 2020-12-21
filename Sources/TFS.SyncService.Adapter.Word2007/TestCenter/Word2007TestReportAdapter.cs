#region Usings
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using AIT.TFS.SyncService.Adapter.Word2007.Properties;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Common.Helper;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Timer = System.Threading.Timer;
#endregion

namespace AIT.TFS.SyncService.Adapter.Word2007.TestCenter
{
    /// <summary>
    /// Class implements <see cref="IWord2007TestReportAdapter"/> - functionality of adapter to create test report in word document.
    /// </summary>
    internal class Word2007TestReportAdapter : IWord2007TestReportAdapter
    {
        #region Private static fields

        private readonly Stack<string> _cursorPositions = new Stack<string>();

        /// <summary>
        /// Holds bookmarks name that should be not deleted in RemoveBookmarks method
        /// </summary>
        private readonly List<string> _storedBookmarks = new List<string>();

        private static readonly AutoResetEvent _wordApplicationEvent = new AutoResetEvent(true);
        private static _Application _wordApplication;
        private static Timer _cleanUpTimer;
        private string _attachmentFolderPath;
        private IConfiguration _configuration;
        private Document _document;
        private PreparationDocumentHelper _preparationDocumentHelper;
        private ITfsTestSuite _testSuite;
        private IList<ITfsTestCaseDetail> _allTestCases;


        /// <summary>
        /// File of opened document. Used only in nested adapter.
        /// </summary>
        private string _file;

        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Word2007TestReportAdapter"/> class.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">The configuration.</param>
        public Word2007TestReportAdapter(Document document, IConfiguration configuration)
        {
            _document = document;
            _configuration = configuration;
            _preparationDocumentHelper = new PreparationDocumentHelper(_document, _wordApplication);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Word2007TestReportAdapter"/> class.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">The configuration.</param>
        /// <param name="testSuite">The test suite</param>
        /// <param name="allTestCases">All test cases related to test suite</param>
        public Word2007TestReportAdapter(Document document, IConfiguration configuration, ITfsTestSuite testSuite, IList<ITfsTestCaseDetail> allTestCases)
        {
            _document = document;
            _configuration = configuration;
            _preparationDocumentHelper = new PreparationDocumentHelper(_document, _wordApplication);
            _testSuite = testSuite;
            _allTestCases = allTestCases;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Word2007TestReportAdapter" /> class.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="file">File to open in adapter.</param>
        /// <param name="configuration">The configuration.</param>
        public Word2007TestReportAdapter(Document document, string file, IConfiguration configuration) : this(document, configuration)
        {
            _file = file;
        }

        #endregion Constructors

        #region Private methods

        /// <summary>
        /// Returns whether the document has the given bookmark
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <returns>
        ///   <c>true</c> if the document has the specified bookmark; otherwise, <c>false</c>.
        /// </returns>
        private bool HasBookmark(string bookmarkName)
        {
            return _document.Bookmarks.Cast<Bookmark>().Any(b => b.Name == bookmarkName);
        }

        /// <summary>
        /// Gets the bookmark with the given name
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <returns>
        /// The bookmark
        /// </returns>
        /// <exception cref="ConfigurationException"></exception>
        private Bookmark GetBookmark(string bookmarkName)
        {
                var bookmark = _document?.Bookmarks?.Cast<Bookmark>().FirstOrDefault( b => b != null && b.Name == bookmarkName);
                if (bookmark == null)
                {
                    throw new ConfigurationException(string.Format(CultureInfo.CurrentCulture, Resources.TestReportError_BookmarkNotFound, bookmarkName));
                }
                return bookmark;
        }


        /// <summary>
        /// The method removes the given bookmark.
        /// </summary>
        /// <param name="bookmarkName">Name of bookmark to remove.</param>
        private void DeleteBookmark(string bookmarkName)
        {
            if (HasBookmark(bookmarkName))
            {
                var bookmark = GetBookmark(bookmarkName);
                bookmark.Delete();
            }
        }

        /// <summary>
        /// The method executes all operations.
        /// </summary>
        /// <param name="operations">List of operations to execute.</param>
        public void ProcessOperations(IList<IConfigurationTestOperation> operations)
        {
            OperationsHelper.ProcessOperations(operations, _document);
        }



        #endregion Private methods

        #region Implementation of IWord2007TestReportAdapter

        /// <summary>
        /// Gets or the bookmarks of the underlying document
        /// </summary>
        public Bookmarks Bookmarks
        {
            get
            {
                return _document.Bookmarks;
            }
        }

        /// <summary>
        /// Replaces an existing bookmark with an error bookmark or creates a new one at the current position
        /// </summary>
        /// <param name="originalBookmarkName">Bookmark to replace</param>
        /// <returns>
        /// The error bookmark name
        /// </returns>
        public string AddErrorBookmark(string originalBookmarkName)
        {
            var bookmarkName = $"c_{Guid.NewGuid().ToString().Replace('-', '_')}";
            if (!string.IsNullOrEmpty(originalBookmarkName) && HasBookmark(originalBookmarkName))
            {
                _document.Bookmarks.Add(bookmarkName, GetBookmark(originalBookmarkName).Range);
            }
            else
            {
                _document.Bookmarks.Add(bookmarkName);
            }

            _storedBookmarks.Add(bookmarkName);
            return bookmarkName;
        }

        /// <summary>
        /// The method saves the actual cursor position.
        /// </summary>
        public void PushCursorPosition()
        {
            var bookmarkName = $"X_{Guid.NewGuid().ToString().Replace('-', '_')}";
            _document.Bookmarks.Add(bookmarkName);
            _cursorPositions.Push(bookmarkName);
        }

        /// <summary>
        /// The method restores previously saved the cursor position.
        /// </summary>
        public void PopCursorPosition()
        {
            if (_cursorPositions.Count <= 0)
            {
                return;
            }
            var bookmarkName = _cursorPositions.Pop();

            if (HasBookmark(bookmarkName))
            {
                var bookmark = GetBookmark(bookmarkName);
                bookmark.Select();
                DeleteBookmark(bookmarkName);
            }
        }

        /// <summary>
        /// The method removes all previously saved cursor positions.
        /// </summary>
        public void ClearCursorPositions()
        {
            while (_cursorPositions.Count > 0)
            {
                PopCursorPosition();
            }
        }

        /// <summary>
        /// The method inserts the content of the word file at the current cursor position.
        /// </summary>
        /// <param name="file">Word file, whose content is to be inserted.</param>
        /// <param name="preOperations">List of operation to process before insert of file.</param>
        /// <param name="postOperations">List of operation to process after insert of file.</param>
        /// <exception cref="ConfigurationException">Thrown if the file does not exist.</exception>
        public void InsertFile(string file, IList<IConfigurationTestOperation> preOperations, IList<IConfigurationTestOperation> postOperations)
        {
            if (!File.Exists(file))
            {
                throw new ConfigurationException(string.Format(CultureInfo.CurrentCulture, Resources.TestReportError_FileNotFound, file));
            }
            ProcessOperations(preOperations);

            try
            {
                _document.ActiveWindow.Selection.InsertFile(file, ConfirmConversions: false);
            }
            catch (Exception e)
            {
                SyncServiceTrace.W(string.Format(CultureInfo.CurrentCulture, Resources.InteropException_ErrorDuringFileInsert, file, e.Message));
                throw new ComInteropException(string.Format(CultureInfo.CurrentCulture, Resources.InteropException_ErrorDuringFileInsert, file, e.Message));
            }

            ProcessOperations(postOperations);
        }

        /// <summary>
        /// The method inserts the content of the word file at the bookmark position.
        /// </summary>
        /// <param name="bookmarkName">Bookmark where is to insert the content of specified file.</param>
        /// <param name="file">Word file, whose content is to be inserted.</param>
        /// <param name="preOperations">List of operation to process before insert of file.</param>
        /// <param name="postOperations">List of operation to process after insert of file.</param>
        /// <exception cref="ConfigurationException">Thrown if the bookmark does not exist.</exception>
        public void InsertFile(string bookmarkName, string file, IList<IConfigurationTestOperation> preOperations, IList<IConfigurationTestOperation> postOperations)
        {
            ProcessOperations(preOperations);
            var bookmark = GetBookmark(bookmarkName);
            bookmark.Select();

            try
            {
                _document.ActiveWindow.Selection.InsertFile(file, ConfirmConversions: false);
                _document.ActiveWindow.Selection.Bookmarks.Add(bookmarkName);
            }
            catch (Exception e)
            {
                SyncServiceTrace.W(string.Format(CultureInfo.CurrentCulture, Resources.InteropException_ErrorDuringFileInsert, file, e.Message));
                throw new ComInteropException(string.Format(CultureInfo.CurrentCulture, Resources.InteropException_ErrorDuringFileInsert, file, e.Message));
            }


            ProcessOperations(postOperations);
        }

        /// <summary>
        /// The method replaces the text of bookmark by specified text.
        /// </summary>
        /// <param name="bookmarkName">Bookmark where is to insert the specified text.</param>
        /// <param name="text">Text which is to insert on the position of the bookmark.</param>
        /// <param name="textFormat">Format of the given text.</param>
        /// <param name="bookmarkToInsertName"></param>
        /// <exception cref="ConfigurationException">Thrown if the bookmark does not exist.</exception>
        public void ReplaceBookmarkText(string bookmarkName, string text, PropertyValueFormat textFormat, string bookmarkToInsertName)
        {
            Guard.ThrowOnArgumentNull(text, "text");

            var bookmark = GetBookmark(bookmarkName);
            bookmark.Select();

            // TODO MIS 6.7.2015: Change to Switch Statement if possible for PropertyValueFormat
            if (textFormat == PropertyValueFormat.PlainText)
            {
                if (!string.IsNullOrEmpty(bookmarkToInsertName))
                {
                    var range = bookmark.Range;
                    range.Text = text;
                    range.Bookmarks.Add(bookmarkToInsertName);
                    StoreOriginalBookmark(bookmarkToInsertName);
                }
                else
                {
                    bookmark.Range.Text = text;
                }
            }
            else
            {
                if (string.IsNullOrEmpty(text))
                {
                    bookmark.Range.Delete();
                }
                else
                {
                    // Editing html fields using vs or browser will not always surround text in tags.
                    // Word appears to only parse strings like "A&amp;B" correctly when they are surrounded with tags.
                    if (!(text.StartsWith("<", StringComparison.Ordinal) && text.EndsWith(">", StringComparison.Ordinal)))
                    {
                        text = $"<span>{text}</span>";
                        SyncServiceTrace.W(Resources.ResultIfTextIsNotHtml);
                    }
                    try
                    {
                        //Used to make numbers of shared steps bold
                        if (PropertyValueFormat.HTMLBold == textFormat)
                        {
                            WordHelper.PasteSpecial(bookmark.Range, text, true);
                        }
                        else
                        {
                            WordHelper.PasteSpecial(bookmark.Range, text);
                        }
                    }
                    catch (COMException)
                    {
                        // If pasting the html text failed, we still have to remove the bookmark text.
                        bookmark.Range.Delete();
                        //bm.Delete();
                    }
                }
            }
        }

        /// <summary>
        /// The method replaces the text of bookmark by specified text.
        /// </summary>
        /// <param name="bookmarkName">Bookmark where is to insert the specified text.</param>
        /// <param name="text">Text which is to insert on the position of the bookmark - text part of hyperlink.</param>
        /// <param name="link">Text which is to insert on the position of the bookmark - link part of hyperlink.</param>
        /// <exception cref="ConfigurationException">Thrown if the bookmark does not exist.</exception>
        public void ReplaceBookmarkHyperlink(string bookmarkName, string text, string link)
        {
            var bookmark = GetBookmark(bookmarkName);
            bookmark.Select();
            bookmark.Range.Hyperlinks.Add(bookmark.Range, link, Type.Missing, link, text);
        }

        /// <summary>
        /// The method inserts the text at the cursor position as heading text with defined level.
        /// </summary>
        /// <param name="text">Text to use.</param>
        /// <param name="headingLevel">Level of the heading. 1 is first level.</param>
        public void InsertHeadingText(string text, int headingLevel)
        {
            if (headingLevel < 0)
            {
                headingLevel = 1;
            }
            object style = WdBuiltinStyle.wdStyleHeading9;
            switch (headingLevel)
            {
                case 1:
                    {
                        style = WdBuiltinStyle.wdStyleHeading1;
                        break;
                    }
                case 2:
                    {
                        style = WdBuiltinStyle.wdStyleHeading2;
                        break;
                    }
                case 3:
                    {
                        style = WdBuiltinStyle.wdStyleHeading3;
                        break;
                    }
                case 4:
                    {
                        style = WdBuiltinStyle.wdStyleHeading4;
                        break;
                    }
                case 5:
                    {
                        style = WdBuiltinStyle.wdStyleHeading5;
                        break;
                    }
                case 6:
                    {
                        style = WdBuiltinStyle.wdStyleHeading6;
                        break;
                    }
                case 7:
                    {
                        style = WdBuiltinStyle.wdStyleHeading7;
                        break;
                    }
                case 8:
                    {
                        style = WdBuiltinStyle.wdStyleHeading8;
                        break;
                    }
                case 9:
                    {
                        style = WdBuiltinStyle.wdStyleHeading9;
                        break;
                    }
            }
            _document.ActiveWindow.Selection.InsertParagraph();
            _document.ActiveWindow.Selection.Move(WdUnits.wdCharacter, -1);
            _document.ActiveWindow.Selection.Select();
            _document.ActiveWindow.Selection.set_Style(style);
            _document.ActiveWindow.Selection.Text = text;
            _document.ActiveWindow.Selection.Move(WdUnits.wdParagraph, 1);
        }

        /// <summary>
        /// The method stores all existing bookmarks and these bookmarks will be not deleted in <see cref="RemoveBookmarks"/>.
        /// </summary>
        public void StoreOriginalBookmarks()
        {
            _storedBookmarks.Clear();
            _storedBookmarks.AddRange(_document.Bookmarks.Cast<Bookmark>().Select(bm => bm.Name).ToList());
        }


        /// <summary>
        /// The method stores an existing bookmark by its name this bookmark will be not deleted in <see cref="RemoveBookmarks"/>.
        /// </summary>
        public void StoreOriginalBookmark(string bookmarkName)
        {
            _storedBookmarks.Add(bookmarkName);
        }


        /// <summary>
        /// The method removes all bookmarks from the document.
        /// </summary>
        public void RemoveBookmarks()
        {
            var list = _document.Bookmarks.Cast<Bookmark>().ToList();
            foreach (var bookmark in list)
            {
                if (_storedBookmarks.Contains(bookmark.Name) || _cursorPositions.Contains(bookmark.Name)) // Don't delete this bookmark
                {
                    continue;
                }
                // Delete this bookmark
                bookmark.Delete();
            }
        }

        /// <summary>
        /// The method creates an adapter for test report opening another file.
        /// </summary>
        /// <param name="file">File to open in adapter.</param>
        /// <returns>Created adapter.</returns>
        public IWord2007TestReportAdapter CreateNestedAdapter(string file)
        {
            if (string.IsNullOrEmpty(file))
            {
                throw new ArgumentNullException("file");
            }
            Document wordDocument;
            try
            {
                // Check word application
                _wordApplicationEvent.WaitOne();
                if (_wordApplication == null)
                {
                    _wordApplication = new Application();
                }
                _wordApplicationEvent.Set();

                // Create document and open file
                wordDocument = _wordApplication.Documents.Open(file, AddToRecentFiles: false, Visible: false);
                wordDocument.Activate();
            }
            catch (Exception e)
            {
                if (_wordApplication != null)
                {
                    _wordApplication.Quit();
                    GC.Collect();
                }

                throw new Exception(String.Format(Resources.TitleOfErrorDuringInitOfWordInstance, file, e.Message));
            }

            return new Word2007TestReportAdapter(wordDocument, file, _configuration)
            {
                ParentAdapter = this
            };
        }

        /// <summary>
        /// Nested adapter created with <see cref="CreateNestedAdapter"/> will be release.
        /// Call of this method for original adapter has no effect.
        /// </summary>
        public void ReleaseNestedAdapter()
        {
            if (!string.IsNullOrEmpty(_file))
            {
                // Close document
                _document.Save();
                // Cast to _Document. In other case the Warning/Error is caused: Ambiguity between method '_Document.Close' and non-method 'DocumentEvent2_Event.Close'
                var document = _document as _Document;
                if (document != null)
                {
                    document.Close();
                }

                // TODO This wont release the adapter if the application crashes in the next second. Also IDisposable may be a much better Idea (unless caching the opened word app is really that much of a time saver) - SER
                // Clean up application
                _wordApplicationEvent.WaitOne();
                if (_wordApplication != null && _wordApplication.Documents != null)
                {
                    if (_wordApplication.Documents.Count == 0)
                    {
                        if (_cleanUpTimer == null)
                        {
                            _cleanUpTimer = new Timer(TimerMethod);
                        }
                        _cleanUpTimer.Change(1000, 0);
                    }
                }
                _wordApplicationEvent.Set();

                // Set fields to default
                _document = null;
                _file = null;
            }
        }

        /// <summary>
        /// Saves the underlying document
        /// </summary>
        public void SaveDocument()
        {
            _document.Save();
        }

        /// <summary>
        /// Gets the path of the underlying document
        /// </summary>
        public string DocumentPath
        {
            get
            {
                return _document.Path;
            }
        }

        /// <summary>
        /// Gets the name of the underlying document
        /// </summary>
        public string DocumentFileName
        {
            get
            {
                return _document.Name;
            }
        }

        /// <summary>
        /// Gets the parent adapter. The parent adapter is the adapter used to create this
        /// adapter if it is a nested adapter or null if it is an adapter working directly
        /// with the report document.
        /// </summary>
        public IWord2007TestReportAdapter ParentAdapter { get; private set; }

        /// <summary>
        /// Gets the absolute path used to save attachments
        /// </summary>
        public string AttachmentFolder
        {
            get
            {
                if (_attachmentFolderPath == null)
                {

                    var attachmentVariable = _document.Variables.Cast<Variable>().FirstOrDefault(x => x.Name.Equals("AITTestReporting_AttachmentFolder"));

                    if (attachmentVariable == null || string.IsNullOrEmpty(attachmentVariable.Value))
                    {
                        if ((_testSuite != null && _configuration.AttachmentFolderMode == AttachmentFolderMode.BasedOnTestSuite) || _configuration.AttachmentFolderMode == AttachmentFolderMode.WithoutGuid)
                        {
                            _attachmentFolderPath = $"Attachments_{Path.GetFileNameWithoutExtension(DocumentFileName)}";
                        }
                        else
                        {
                            _attachmentFolderPath = $"Attachments_{Path.GetFileNameWithoutExtension(DocumentFileName)}_{Guid.NewGuid().ToString()}";
                        }
                        _document.Variables.Add("AITTestReporting_AttachmentFolder", _attachmentFolderPath);
                    }
                    else
                    {
                        _attachmentFolderPath = attachmentVariable.Value;
                    }
                }

                return _attachmentFolderPath;
            }
        }

        /// <summary>
        /// Local file name of attachment
        /// </summary>
        /// <param name="attachmentName">Name of attachment</param>
        /// <param name="testCase">Tast case that contains attachment</param>
        /// <returns></returns>
        public string AttachmentFileName(string attachmentName, ITfsTestCaseDetail testCase)
        {
            var attachementFileName = string.Empty;
            foreach (var testCaseItem in _allTestCases)
            {
                if(testCaseItem.OriginalTestCase.Attachments.Count > 0 && testCaseItem.OriginalTestCase.Attachments.Where(x => x.Name.Equals(attachmentName)) != null)
                {
                    attachementFileName = $"\\{testCase.Title}_({testCase.Id})";
                }
            }

            return attachementFileName;
        }

        /// <summary>
        /// Gets or sets the hyperlink base set in the document properties
        /// </summary>
        /// <remarks>
        /// This property takes a non-negligible time to access document properties. Make sure to cache it when using in a loop.
        /// </remarks>
        public string HyperlinkBase
        {
            get
            {
                // Okay, according to this article http://support.microsoft.com/kb/303296/en-us
                // you have to use reflection because you somehow cannot cast the builtindocumentproperties
                // into the interface IBuiltInDocumentProperties

                // Attention: The next line takes LOTS of actual time. Make sure to cache HyperlinkBase when using in a loop
                var wtf = _document.BuiltInDocumentProperties;
                var wtfType = wtf.GetType();
                var hlbProperty = wtfType.InvokeMember("Item",
                                                       BindingFlags.Default | BindingFlags.GetProperty,
                                                       null,
                                                       wtf,
                                                       new object[]
                                                           {
                                                               WdBuiltInProperty.wdPropertyHyperlinkBase
                                                           }, CultureInfo.InvariantCulture);

                var hlbPropertyType = hlbProperty.GetType();
                var value = hlbPropertyType.InvokeMember("Value",
                                                         BindingFlags.Default | BindingFlags.GetProperty,
                                                         null,
                                                         hlbProperty,
                                                         new object[]
                                                             {
                                                             }, CultureInfo.InvariantCulture);



                return value == null ? null : value.ToString();
            }
            set
            {
                var wtf = _document.BuiltInDocumentProperties;
                var wtfType = wtf.GetType();
                wtfType.InvokeMember("Item",
                                     BindingFlags.Default | BindingFlags.SetProperty,
                                     null,
                                     wtf,
                                     new object[]
                                         {
                                             WdBuiltInProperty.wdPropertyHyperlinkBase, value
                                         }, CultureInfo.InvariantCulture);

                // Interestingly, this does not happen automatically...
                _document.Saved = false;
            }
        }

        /// <summary>
        /// Checks the hyperlink base is set to the attachment folder.
        /// Checks if the attachment folder is in the document location.
        /// Checks if document contains link to files that are not in the attachment folder.
        /// Checks if attachment folder contains files that are not linked to in the document.
        /// </summary>
        public void CheckAttachmentConsistency()
        {
            if (_configuration.ConfigurationTest.SetHyperlinkBase)
            {
                if (HyperlinkBase == null || !HyperlinkBase.Equals(_document.Path))
                {
                    if (HasInvalidAttachmentLinks())
                    {
                            // Document has moved, attachment folder has moved but hyperlink base still points to old location.
                       if (Directory.Exists(Path.Combine(_document.Path, AttachmentFolder)) &&
                                DialogResult.Yes ==
                                                    MessageBox.Show(Resources.Attachment_AdjustHyperlinkBaseText,
                                                    Resources.Attachment_AdjustHyperlinkBaseTitle,
                                                    MessageBoxButtons.YesNo,
                                                    MessageBoxIcon.Question,
                                                    MessageBoxDefaultButton.Button1,
                                                    0))
                        {
                                HyperlinkBase = _document.Path;
                                _document.Saved = false;
                        }
                    }
                    else
                    {
                        // Document moved but attachment folder at old location. Move folder, adjust hyperlink base
                        if (HyperlinkBase != null && Directory.Exists(Path.Combine(HyperlinkBase, AttachmentFolder)) &&
                            DialogResult.Yes ==
                                                MessageBox.Show(string.Format(CultureInfo.CurrentUICulture, 
                                                Resources.Attachment_MoveAttachmentFolderText, AttachmentFolder, HyperlinkBase),
                                                Resources.Attachment_AdjustHyperlinkBaseTitle,
                                                MessageBoxButtons.YesNo,
                                                MessageBoxIcon.Question,
                                                MessageBoxDefaultButton.Button1,
                                                0))
                        {
                            Directory.Move(Path.Combine(HyperlinkBase, AttachmentFolder), Path.Combine(_document.Path, AttachmentFolder));
                            HyperlinkBase = _document.Path;
                            _document.Saved = false;
                        }
                }
            }
        } 

            // if there are still broken links, warn and suggest a regeneration
            if (HasInvalidAttachmentLinks())
            {
                MessageBox.Show(Resources.Attachment_InconsistentFolderStructureText,
                                Resources.Attachment_InconsistentFolderStructureTitle,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                0);
            }

            // get all files that have no hyperlink pointing at them
            var hyperlinkBase = string.IsNullOrEmpty(HyperlinkBase) ? _document.Path : HyperlinkBase;
            if (Directory.Exists(hyperlinkBase))
            {
                var hyperlinks = _document.Hyperlinks.Cast<Hyperlink>().ToList();
                var attachmentFolder = Path.Combine(hyperlinkBase, AttachmentFolder);

                if (Directory.Exists(attachmentFolder))
                {
                    var orphanedFiles = Directory.GetFiles(attachmentFolder).Where(file => !hyperlinks.Any(h => file.Equals(Path.Combine(hyperlinkBase, h.Address)))).ToList();

                      if (orphanedFiles.Any() && DialogResult.Yes ==
                            MessageBox.Show(Resources.Attachment_DeleteAttachmentFolderText,
                                            Resources.Attachment_DeleteAttachmentFolderTitle,
                                            MessageBoxButtons.YesNo,
                                            MessageBoxIcon.Question,
                                            MessageBoxDefaultButton.Button1,
                                            0))
                        {
                            foreach (var file in orphanedFiles)
                            {
                                File.Delete(file);
                            }
                        }
                }
            }
        }


        private bool HasInvalidAttachmentLinks()
        {
            // check if there are broken links to local files
            var links = _document.Hyperlinks;
            var oldSaved = _document.Saved;
            var hyperlinkBase = string.IsNullOrEmpty(HyperlinkBase) ? _document.Path : HyperlinkBase;
            var hasBrokenLinks =
                links.Cast<Hyperlink>()
                     .Any(x => x.Address != null && !x.Address.StartsWith("http", StringComparison.OrdinalIgnoreCase) && !File.Exists(Path.Combine(hyperlinkBase, x.Address)));

            // it seems querying the hyperlinks causes the document to become dirty
            _document.Saved = oldSaved;
            return hasBrokenLinks;
        }

        private static void TimerMethod(object state)
        {
            _wordApplicationEvent.WaitOne();
            _cleanUpTimer.Dispose();
            _cleanUpTimer = null;
            if (_wordApplication.Documents.Count == 0)
            {
                _wordApplication.Quit();
                _wordApplication = null;
                GC.Collect();
            }
            _wordApplicationEvent.Set();
        }

        #endregion Implementation of IWord2007TestReportAdapter

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
    }
}