using Microsoft.Office.Interop.Word;
using System;

namespace AIT.TFS.SyncService.Contracts.Model
{
    /// <summary>
    /// Event arguments for the document open event.
    /// </summary>
    public class DocumentOpenEventArgs : EventArgs
    {
        /// <summary>
        /// The opened document.
        /// </summary>
        public Document Document
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentOpenEventArgs"/> class.
        /// </summary>
        /// <param name="document">The opened document.</param>
        public DocumentOpenEventArgs(Document document)
        {
            Document = document;
        }
    }
}
