using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Exceptions;
using Microsoft.Office.Interop.Word;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of adapter to create test report in word document.
    /// </summary>
    public interface IWord2007TestReportAdapter: IPreparationDocument
    {
        /// <summary>
        /// The method saves the actual cursor position.
        /// </summary>
        void PushCursorPosition();

        /// <summary>
        /// The method restores previously saved the cursor position.
        /// </summary>
        void PopCursorPosition();

        /// <summary>
        /// The method removes all previously saved cursor positions.
        /// </summary>
        void ClearCursorPositions();

        /// <summary>
        /// The method inserts the content of the word file at the current cursor position.
        /// </summary>
        /// <param name="file">Word file, whose content is to be inserted.</param>
        /// <param name="preOperations">List of operation to process before insert of file.</param>
        /// <param name="postOperations">List of operation to process after insert of file.</param>
        /// <exception cref="ConfigurationException">Thrown if the bookmark does not exist.</exception>
        void InsertFile(string file, IList<IConfigurationTestOperation> preOperations, IList<IConfigurationTestOperation> postOperations);

        /// <summary>
        /// The method inserts the content of the word file at the bookmark position.
        /// </summary>
        /// <param name="bookmarkName">Bookmark where is to insert the content of specified file.</param>
        /// <param name="file">Word file, whose content is to be inserted.</param>
        /// <param name="preOperations">List of operation to process before insert of file.</param>
        /// <param name="postOperations">List of operation to process after insert of file.</param>
        /// <exception cref="ConfigurationException">Thrown if the bookmark does not exist.</exception>
        void InsertFile(string bookmarkName, string file, IList<IConfigurationTestOperation> preOperations, IList<IConfigurationTestOperation> postOperations);

        /// <summary>
        /// The method replaces the text of bookmark by specified text.
        /// </summary>
        /// <param name="bookmarkName">Bookmark where is to insert the specified text.</param>
        /// <param name="text">Text which is to insert on the position of the bookmark.</param>
        /// <param name="textFormat">Format of the given text.</param>
        /// <param name="additionalBookmarkName">Name of an additional Bookmark that should be inserted after the update</param>
        /// <exception cref="ConfigurationException">Thrown if the bookmark does not exist.</exception>
        void ReplaceBookmarkText(string bookmarkName, string text, PropertyValueFormat textFormat,string additionalBookmarkName);

        /// <summary>
        /// The method replaces the text of bookmark by specified text.
        /// </summary>
        /// <param name="bookmarkName">Bookmark where is to insert the specified text.</param>
        /// <param name="text">Text which is to insert on the position of the bookmark - text part of hyperlink.</param>
        /// <param name="link">Text which is to insert on the position of the bookmark - link part of hyperlink.</param>
        /// <exception cref="ConfigurationException">Thrown if the bookmark does not exist.</exception>
        void ReplaceBookmarkHyperlink(string bookmarkName, string text, string link);

        /// <summary>
        /// The method inserts the text at the cursor position as heading text with defined level.
        /// </summary>
        /// <param name="text">Text to use.</param>
        /// <param name="headingLevel">Level of the heading.</param>
        void InsertHeadingText(string text, int headingLevel);

        /// <summary>
        /// The method stores all existing bookmarks and these bookmarks will be not deleted in <see cref="RemoveBookmarks"/>.
        /// </summary>
        void StoreOriginalBookmarks();

        /// <summary>
        /// The method removes all bookmarks from the document.
        /// </summary>
        void RemoveBookmarks();

        /// <summary>
        /// The method creates an adapter for test report opening another file.
        /// </summary>
        /// <param name="file">File to open in adapter.</param>
        /// <returns>Created adapter.</returns>
        IWord2007TestReportAdapter CreateNestedAdapter(string file);

        /// <summary>
        /// Nested adapter created with <see cref="CreateNestedAdapter"/> will be release.
        /// Call of this method for original adapter has no effect.
        /// </summary>
        void ReleaseNestedAdapter();

        /// <summary>
        /// Saves the underlying document
        /// </summary>
        void SaveDocument();

        /// <summary>
        /// Gets the path of the underlying document
        /// </summary>
        string DocumentPath
        {
            get;
        }

        /// <summary>
        /// Gets the name of the underlying document
        /// </summary>
        string DocumentFileName
        {
            get;
        }

        /// <summary>
        /// Gets or the bookmarks of the underlying document
        /// </summary>
        Bookmarks Bookmarks
        {
            get;
        }

        /// <summary>
        /// Replaces an existing bookmark with an error bookmark or creates a new one at the current position
        /// </summary>
        /// <param name="originalBookmarkName">Bookmark to replace</param>
        /// <returns>The error bookmark name</returns>
        string AddErrorBookmark(string originalBookmarkName);

        /// <summary>
        /// Gets the parent adapter. The parent adapter is the adapter used to create this
        /// adapter if it is a nested adapter or null if it is an adapter working directly
        /// with the report document.
        /// </summary>
        IWord2007TestReportAdapter ParentAdapter { get; }

        /// <summary>
        /// Gets or sets the hyperlink base set in the document properties
        /// </summary>
        string HyperlinkBase { get; set; }

        /// <summary>
        /// Gets the absolute path used to save attachments
        /// </summary>
        string AttachmentFolder { get; }

        /// <summary>
        /// Gets attachment file name.
        /// </summary>
        string AttachmentFileName(string attachmentName, ITfsTestCaseDetail testCase);

        /// <summary>
        /// Checks the hyperlink base is set to the attachment folder.
        /// Checks if the attachment folder is in the document location.
        /// Checks if document contains link to files that are not in the attachment folder.
        /// Checks if attachment folder contains files that are not linked to in the document.
        /// </summary>
        void CheckAttachmentConsistency();


        /// <summary>
        /// Process any operations
        /// </summary>
        /// <param name="operations"></param>
        void ProcessOperations(IList<IConfigurationTestOperation> operations);

    }
}
