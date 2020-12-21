using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    /// Interface defines configuration for test template.
    /// </summary>
    public interface IConfigurationTestTemplate
    {
        /// <summary>
        /// Gets the name of template.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets the template name for header of block.
        /// </summary>
        string HeaderTemplate { get; }

        /// <summary>
        /// Gets the file name of template.
        /// </summary>
        string FileName { get; }

        /// <summary>
        /// Gets the list of pre-operations.
        /// </summary>
        IList<IConfigurationTestOperation> PreOperations { get; }

        /// <summary>
        /// Gets the list of post-operations.
        /// </summary>
        IList<IConfigurationTestOperation> PostOperations { get; }

        /// <summary>
        /// Gets the list of conditions in one template.
        /// </summary>
        IList<IConfigurationTestCondition> Conditions { get; }

        /// <summary>
        /// Gets the list of replacements in one template.
        /// </summary>
        IList<IConfigurationTestReplacement> Replacements { get; }
    }
}
