namespace AIT.TFS.SyncService.Contracts.TemplateManager
{
    /// <summary>
    /// Enumeration defines the state of one template in template bundle.
    /// </summary>
    public enum TemplateState
    {
        /// <summary>
        /// The template is not initialized.
        /// </summary>
        NotInitialized = 0,

        /// <summary>
        /// The template is initialized and available.
        /// </summary>
        Available,

        /// <summary>
        /// The template is initialized and disabled.
        /// </summary>
        Disabled,

        /// <summary>
        /// The template file (w2t) is faulty.
        /// </summary>
        ErrorInTemplate,
    }
}
