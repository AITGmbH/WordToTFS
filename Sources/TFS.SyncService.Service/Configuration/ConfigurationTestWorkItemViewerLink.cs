using System;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationTestWorkItemViewerLink"/>.
    /// </summary>
    public class ConfigurationTestWorkItemViewerLink : IConfigurationTestWorkItemViewerLink
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestWorkItemViewerLink"/> class.
        /// </summary>
        /// <param name="workItemViewerLink">Associated configuration.</param>
        public ConfigurationTestWorkItemViewerLink(WorkItemViewerLink workItemViewerLink)
        {
            if (workItemViewerLink == null)
                throw new ArgumentNullException("workItemViewerLink");
            Id = workItemViewerLink.Id;
            Revision = workItemViewerLink.Revision;
            AutoText = workItemViewerLink.AutoText;
            Format = workItemViewerLink.Format;
        }

        #endregion Constructors

        #region Implementation of IConfigurationTestWorkItemViewerLink

        /// <summary>
        /// Gets or sets the property to evaluate and gets the id of the work item.
        /// </summary>
        public string Id { get; private set; }

        /// <summary>
        /// Gets or sets the property to evaluate and gets the revision of the work item.
        /// </summary>
        public string Revision { get; private set; }

        /// <summary>
        /// Gets the flag if the configured text for Hyperlink will be suppressed and text 'id - title' will be generated.
        /// </summary>
        public bool AutoText { get; private set; }

        /// <summary>
        /// Gets the string replacement formating of the <see cref="IConfigurationTestWorkItemViewerLink"/>
        /// </summary>
        public string Format { get; private set; }

        #endregion Implementation of IConfigurationTestWorkItemViewerLink
    }
}
