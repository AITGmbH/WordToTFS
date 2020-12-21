#region Usings
using Microsoft.Office.Interop.Word;
#endregion

namespace AIT.TFS.SyncService.Common.Helper
{
    /// <summary>
    /// This is helper class which is used for preparation document for long-term operations
    /// </summary>
    public class PreparationDocumentHelper
    {
        #region Fields
        private Document _document;
        private _Application _wordApplication;

        private bool _wasPaginationOn;
        private bool _wasShowGrammaticalErrorsOn;
        private bool _wasShowSpellingErrorsOn;
        private bool _wasWordApplicationPaginationOn;
        private bool _wasCheckGrammarAsYouTypeOn;
        private bool _wasCheckSpellingAsYouTypeOn;
        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="PreparationDocumentHelper"/> class.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="wordApplication"></param>
        public PreparationDocumentHelper(Document document, _Application wordApplication)
        {
            _document = document;
            _wordApplication = wordApplication;
        }
        #endregion

        #region Public methods

        /// <summary>
        /// Prepare the document for the long-term operation. 
        /// Disable Spellchecking and Pagination. This is necessary, because Word Add-ins are not supposed to create big documents in a self running manner
        /// </summary>
        public void PrepareDocumentForLongTermOperation()
        {
            //Temporary storage of properties
            _wasPaginationOn = _document.Application.Options.Pagination;
            _wasShowGrammaticalErrorsOn = _document.ShowGrammaticalErrors;
            _wasShowSpellingErrorsOn = _document.ShowSpellingErrors;
            if (_wordApplication != null)
            {
                _wasWordApplicationPaginationOn = _wordApplication.Application.Options.Pagination;
                _wasCheckGrammarAsYouTypeOn = _wordApplication.Application.Options.CheckGrammarAsYouType;
                _wasCheckSpellingAsYouTypeOn = _wordApplication.Application.Options.CheckSpellingAsYouType;
            }

            SetDocumentForLongTermOperation(false);
        }

        /// <summary>
        /// Undo the changes to the document.
        /// Enable Spellchecking and Pagination.
        /// </summary>
        public void UndoPreparationsDocumentForLongTermOperation()
        {
            _document.Application.Options.Pagination = _wasPaginationOn;
            _document.ShowGrammaticalErrors = _wasShowGrammaticalErrorsOn;
            _document.ShowSpellingErrors = _wasShowSpellingErrorsOn;

            if (_wordApplication != null)
            {
                _wordApplication.Application.Options.Pagination = _wasWordApplicationPaginationOn;
                _wordApplication.Application.Options.CheckGrammarAsYouType = _wasCheckGrammarAsYouTypeOn;
                _wordApplication.Application.Options.CheckSpellingAsYouType = _wasCheckSpellingAsYouTypeOn;

            }
        }
        #endregion

        #region Private methods
        private void SetDocumentForLongTermOperation(bool value)
        {
            _document.Application.Options.Pagination = value;
            _document.ShowGrammaticalErrors = value;
            _document.ShowSpellingErrors = value;

            if (_wordApplication != null)
            {
                _wordApplication.Application.Options.Pagination = value;
                _wordApplication.Application.Options.CheckGrammarAsYouType = value;
                _wordApplication.Application.Options.CheckSpellingAsYouType = value;

            }
        }
        #endregion
    }
}
