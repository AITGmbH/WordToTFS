using System;

namespace AIT.TFS.SyncService.Contracts.Model
{
    /// <summary>
    /// Interface defines the functionality of the model for add-in.
    /// </summary>
    public interface ISyncServiceModel
    {
        /// <summary>
        /// The method gets the model of the given word document.
        /// </summary>
        /// <param name="wordDocument">Word document to get the <see cref="ISyncServiceDocumentModel"/> for.</param>
        /// <returns>Required <see cref="ISyncServiceDocumentModel"/>.</returns>
        ISyncServiceDocumentModel GetModel(object wordDocument);

        /// <summary>
        /// The method checks whether all documents for stored <see cref="ISyncServiceDocumentModel"/> still exist.
        /// </summary>
        void CleanUpModelList();

        /// <summary>
        /// Occurs when the active document changed. No document may be active if no document opened.
        /// </summary>
        event EventHandler<EventArgs> ActiveDocumentChanged;

        /// <summary>
        /// Occurs when a document is opened.
        /// </summary>
        event EventHandler<DocumentOpenEventArgs> DocumentOpen;
    }
}