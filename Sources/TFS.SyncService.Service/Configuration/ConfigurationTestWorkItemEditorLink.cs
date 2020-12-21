using System;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationTestWorkItemEditorLink"/>.
    /// </summary>
    public class ConfigurationTestWorkItemEditorLink : IConfigurationTestWorkItemEditorLink
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestWorkItemEditorLink"/> class.
        /// </summary>
        /// <param name="workItemEditorLink">Associated configuration.</param>
        public ConfigurationTestWorkItemEditorLink(WorkItemEditorLink workItemEditorLink)
        {
            if (workItemEditorLink == null)
                throw new ArgumentNullException("workItemEditorLink");
            Uri = workItemEditorLink.Uri;
        }

        #endregion Constructors

        #region Implementation of IConfigurationTestWorkItemEditorLink

        /// <summary>
        /// Gets or sets the property to evaluate and gets the artifact uri of the work item editor.
        /// </summary>
        public string Uri { get; private set; }

        #endregion Implementation of IConfigurationTestWorkItemEditorLink
    }
}
