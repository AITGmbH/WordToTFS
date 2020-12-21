namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Enumeration defines all supported operation before / insert of template.
    /// </summary>
    public enum OperationType
    {
        /// <summary>
        /// No operation
        /// </summary>
        None = 0,

        /// <summary>
        /// In target document is new paragraph inserted on the cursor position.
        /// </summary>
        InsertParagraph,

        /// <summary>
        /// In target document is cursor moved to front of document.
        /// </summary>
        MoveCursorToStart,

        /// <summary>
        /// In target document is cursor moved to end of document.
        /// </summary>
        MoveCursorToEnd,

        /// <summary>
        /// In target document is deleted one character on the left side of cursor.
        /// </summary>
        DeleteCharacterLeft,

        /// <summary>
        /// In target document is deleted one character on the left side of cursor.
        /// </summary>
        DeleteCharacterRight,

        /// <summary>
        /// In target document is cursor moved one character left.
        /// </summary>
        MoveCursorToLeft,

        /// <summary>
        /// In target document is cursor moved one character left.
        /// </summary>
        MoveCursorToRight,

        /// <summary>
        /// Insert a new page
        /// </summary>
        InsertNewPage,

        /// <summary>
        /// Refresh all fields in Document
        /// </summary>
        RefreshAllFieldsInDocument
    }
}
