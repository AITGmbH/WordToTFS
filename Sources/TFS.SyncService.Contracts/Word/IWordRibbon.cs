namespace AIT.TFS.SyncService.Contracts.Word
{
    using AIT.TFS.SyncService.Contracts.Model;

    /// <summary>
    /// Interface for the AIT WordToTFS word ribbon.
    /// </summary>
    public interface IWordRibbon
    {
        /// <summary>
        /// The method sets / resets all required members to appropriate state.
        /// </summary>
        /// <param name="documentModel">Model of the document where will be the operation started.</param>
        void ResetBeforeOperation(ISyncServiceDocumentModel documentModel);

        /// <summary>
        /// Method sets the state of all controls to desired state that depends on the model - active document.
        /// </summary>
        /// <param name="documentModel">Associated model of document to perform the action for.</param>
        void EnableAllControls(ISyncServiceDocumentModel documentModel);

        /// <summary>
        /// Disables all controls.
        /// </summary>
        /// <param name="documentModel">Associated model of document to perform the action for.</param>
        void DisableAllControls(ISyncServiceDocumentModel documentModel);
    }
}
