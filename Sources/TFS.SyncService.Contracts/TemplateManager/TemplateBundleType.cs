namespace AIT.TFS.SyncService.Contracts.TemplateManager
{
    /// <summary>
    /// Enumeration defines all possible types of template bundle - collection of w2t files in one folder.
    /// </summary>
    public enum TemplateBundleType
    {
        /// <summary>
        /// The template has unknown origin.
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// The template bundle is from local disk.
        /// </summary>
        LocalBundle,

        /// <summary>
        /// The template bundle is from shared folder in network.
        /// </summary>
        UncBundle,

        /// <summary>
        /// The template bundle is from web resource.
        /// </summary>
        WebBundle,
    }
}
