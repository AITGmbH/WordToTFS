#region Usings
using System.Windows.Forms;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Properties;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using Application = Microsoft.Office.Interop.Word.Application;
#endregion

namespace AIT.TFS.SyncService.Model
{
    /// <summary>
    /// Class implements the interface <see cref="ISyncServiceModel"/> - model for add-in.
    /// </summary>
    public class SyncServiceModel : ISyncServiceModel
    {
        #region Private fields
        
        private readonly List<ISyncServiceDocumentModel> _documentModels = new List<ISyncServiceDocumentModel>();
        
        #endregion Private fields

        #region Constructors
        
        /// <summary>
        /// Initializes a new instance of the <see cref="SyncServiceModel"/> class.
        /// </summary>
        /// <param name="wordApplication">Associated word application</param>
        public SyncServiceModel(Application wordApplication)
        {
            if (wordApplication == null)
                throw new ArgumentNullException("wordApplication");
            WordApplication = wordApplication;
            WordApplication.DocumentChange += OnActiveDocumentChanged;
            WordApplication.DocumentOpen += WordApplicationOnDocumentOpen;
            wordApplication.DocumentBeforeSave += WordApplicationOnDocumentBeforeSave;
        }

        /// <summary>
        /// Warn that "save as" is not really recommended for report documents with links to local files.
        /// Those links are relative to the document hyperlink base which points to the document location
        /// so moving the document in a different directory breaks those links.
        /// </summary>
        private void WordApplicationOnDocumentBeforeSave(Document doc, ref bool saveAsUi, ref bool cancel)
        {
            var model = GetModel(doc);
            if (model.TestReportGenerated && saveAsUi)
            {
                if (doc.Hyperlinks.Cast<Hyperlink>().Any(x => x.Address.StartsWith("http", StringComparison.OrdinalIgnoreCase) == false))
                {
                    MessageBox.Show(Resources.Attachment_SaveAsWarning, Resources.MessageBox_Title, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
                }
            }
        }

        #endregion Constructors

        #region Public properties
        
        /// <summary>
        /// Gets the word application.
        /// </summary>
        public Application WordApplication { get; private set; }
        
        #endregion Public properties

        #region Public methods

        /// <summary>
        /// The method checks whether all documents for stored <see cref="ISyncServiceDocumentModel"/> still exist.
        /// </summary>
        public void CleanUpModelList()
        {
            lock (_documentModels)
            {
                var removed = true;
                while (removed)
                {
                    removed = false;
                    foreach (var model in _documentModels)
                    {
                        if (!WordApplication.Documents.Cast<object>().Any(document => model.WordDocument == document))
                        {
                            _documentModels.Remove(model);
                            // Don't iterate to next items. It will be checked next time.
                            removed = true;
                            break;
                        }
                    }
                }
            }
        }

        #endregion Public methods

        #region Private methods

        /// <summary>
        /// Check attachment links for generated reporting documents
        /// </summary>
        private void WordApplicationOnDocumentOpen(Document document)
        {
            if (document != null)
            {
                var model = GetModel(document);
                if (model.TestReportGenerated)
                {
                    var adapter = SyncServiceFactory.CreateWord2007TestReportAdapter(document, model.Configuration);
                    adapter.CheckAttachmentConsistency();
                }
            }

            if (DocumentOpen != null)
                DocumentOpen(this, new DocumentOpenEventArgs(document));
        }

        /// <summary>
        /// Called when a new document is created, when an existing document is opened, or when another document is made the active document.
        /// </summary>
        protected virtual void OnActiveDocumentChanged()
        {
            if (ActiveDocumentChanged != null)
                ActiveDocumentChanged(this, EventArgs.Empty);
        }

        #endregion Private methods

        #region Implementation of ISyncServiceModel

        /// <summary>
        /// The method gets the model of the given word document.
        /// </summary>
        /// <param name="wordDocument">Word document to get the <see cref="ISyncServiceDocumentModel"/> for.</param>
        /// <returns>Required <see cref="ISyncServiceDocumentModel"/>.</returns>
        public ISyncServiceDocumentModel GetModel(object wordDocument)
        {
            var document = wordDocument as Document;
            if (document != null)
            {
                lock (_documentModels)
                {
                    var modelToReturn = _documentModels.FirstOrDefault(model => model.WordDocument == document);
                    if (modelToReturn != null) return modelToReturn;
                    CleanUpModelList();

                    // insert new model and return it
                    _documentModels.Add(new SyncServiceDocumentModel(wordDocument as Document));
                    modelToReturn = _documentModels.FirstOrDefault(model => model.WordDocument == document);
                    if (modelToReturn != null) return modelToReturn;
                }
            }

            // TODO Actually we should throw an error and remove all those null checks. Not having a model is really bad and if we only want to hide the exception, change the UnhandledException-Handler
            SyncServiceTrace.E("No model found for document or document is null");
            return null;
        }

        /// <summary>
        /// Occurs when the active document changed. No document may be active if no document opened.
        /// </summary>
        public event EventHandler<EventArgs> ActiveDocumentChanged;

        /// <summary>
        /// Occurs when a document is opened.
        /// </summary>
        public event EventHandler<DocumentOpenEventArgs> DocumentOpen;

        #endregion Implementation of ISyncServiceModel

    }
}
