namespace AIT.TFS.SyncService.Contracts.TemplateManager
{
    using System;

    /// <summary>
    /// Enumeration defines all possible states of one template bundle.
    /// </summary>
    [Flags]
    public enum TemplateBundleStates
    {
        /// <summary>
        /// All flags cleared
        /// </summary>
        None = 0x00,

        /// <summary>
        /// Template bundle is available.
        /// </summary>
        TemplateBundleAvailable = 0x02,

        /// <summary>
        /// Template bundle is cached - origin of the template bundle is not available.
        /// </summary>
        TemplateBundleCached = 0x04,

        /// <summary>
        /// Template bundle contains at least one faulty template.
        /// </summary>
        AtLeastOneTemplateIsFaulty = 0x08,

        /// <summary>
        /// Template bundle contains at least one disabled template.
        /// </summary>
        AtLeastOneTemplateDisabled = 0x10,
    }
}
