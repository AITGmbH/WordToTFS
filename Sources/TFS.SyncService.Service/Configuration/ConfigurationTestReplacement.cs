using System;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
	using AIT.TFS.SyncService.Contracts.Enums;

	/// <summary>
    /// The class implements <see cref="IConfigurationTestReplacement"/>.
    /// </summary>
    public class ConfigurationTestReplacement : IConfigurationTestReplacement
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestReplacement"/> class
        /// </summary>
        /// <param name="testReplace">Associated configuration from w2t file.</param>
        public ConfigurationTestReplacement(ReplacementConfiguration testReplace)
        {
            if (testReplace == null)
                throw new ArgumentNullException("testReplace");
            WordBookmark = testReplace.WordBookmark;
            Bookmark = testReplace.Bookmark;
			VariableName = testReplace.VariableName;
			FieldValueType = testReplace.FieldValueType;
			PropertyToEvaluate = testReplace.PropertyToEvaluate;
            ValueType = testReplace.ValueType;
            LinkedTemplate = testReplace.LinkedTemplate;
            Parameters = testReplace.Parameters;
            ResolveResolutionState = testReplace.ResolveResolutionState;
            if (testReplace.WorkItemEditorLink != null)
                WorkItemEditorLink = new ConfigurationTestWorkItemEditorLink(testReplace.WorkItemEditorLink);
            if (testReplace.WorkItemViewerLink != null)
                WorkItemViewerLink = new ConfigurationTestWorkItemViewerLink(testReplace.WorkItemViewerLink);
            if (testReplace.BuildViewerLink != null)
                BuildViewerLink = new ConfigurationTestBuildViewerLink(testReplace.BuildViewerLink);
            if (testReplace.UriLink != null)
                UriLink = new ConfigurationTestUriLink(testReplace.UriLink);
            if (testReplace.AttachmentLink != null)
                AttachmentLink = new ConfigurationAttachmentLink(testReplace.AttachmentLink);
        }

		#endregion Constructors

        #region Implementation of IConfigurationTestTemplate

        /// <summary>
        /// Gets the name of bookmark to replace.
        /// </summary>
        public string Bookmark { get; private set; }

		/// <summary>
		/// Gets the variable name if a variable should be used for replacement.
		/// </summary>
		public string VariableName { get; private set; }

        /// <summary>
        /// Gets the <see cref="IConfigurationTestReplacement.FieldValueType"/> for the replacment.
        /// </summary>
        public FieldValueType FieldValueType { get; private set; }

		/// <summary>
        /// Gets the property to evaluate and gets the value for replacement - used as text to show.
        /// </summary>
        public string PropertyToEvaluate { get; private set; }

        /// <summary>
        /// Gets the optinal parameters for a property
        /// </summary>
        public string Parameters
        {
            get;
            private set;
        }
        /// <summary>
        /// Gets the type of value provided by evaluated of property.
        /// </summary>
        public PropertyValueFormat ValueType { get; private set; }

        /// <summary>
        /// Gets the linked template to use for enumerable properties by replacement.
        /// </summary>
        /// <remarks>
        /// Not null value returns either no of these properties or only one property:
        /// <list type="bullet">
        /// <item><description><see cref="LinkedTemplate"/></description></item>
        /// <item><description><see cref="WorkItemEditorLink"/></description></item>
        /// <item><description><see cref="WorkItemViewerLink"/></description></item>
        /// <item><description><see cref="BuildViewerLink"/></description></item>
        /// <item><description><see cref="UriLink"/></description></item>
        /// </list>
        /// </remarks>
        public string LinkedTemplate { get; private set; }

        /// <summary>
        /// Gets the configuration for work item editor link.
        /// </summary>
        /// <remarks>
        /// Not null value returns either no of these properties or only one property:
        /// <list type="bullet">
        /// <item><description><see cref="LinkedTemplate"/></description></item>
        /// <item><description><see cref="WorkItemEditorLink"/></description></item>
        /// <item><description><see cref="WorkItemViewerLink"/></description></item>
        /// <item><description><see cref="BuildViewerLink"/></description></item>
        /// <item><description><see cref="UriLink"/></description></item>
        /// </list>
        /// </remarks>
        public IConfigurationTestWorkItemEditorLink WorkItemEditorLink { get; private set; }

        /// <summary>
        /// Gets the configuration for work item viewer link.
        /// </summary>
        /// <remarks>
        /// Not null value returns either no of these properties or only one property:
        /// <list type="bullet">
        /// <item><description><see cref="LinkedTemplate"/></description></item>
        /// <item><description><see cref="WorkItemEditorLink"/></description></item>
        /// <item><description><see cref="WorkItemViewerLink"/></description></item>
        /// <item><description><see cref="BuildViewerLink"/></description></item>
        /// <item><description><see cref="UriLink"/></description></item>
        /// </list>
        /// </remarks>
        public IConfigurationTestWorkItemViewerLink WorkItemViewerLink { get; private set; }

        /// <summary>
        /// Gets the configuration for build viewer link.
        /// </summary>
        /// <remarks>
        /// Not null value returns either no of these properties or only one property:
        /// <list type="bullet">
        /// <item><description><see cref="LinkedTemplate"/></description></item>
        /// <item><description><see cref="WorkItemEditorLink"/></description></item>
        /// <item><description><see cref="WorkItemViewerLink"/></description></item>
        /// <item><description><see cref="BuildViewerLink"/></description></item>
        /// <item><description><see cref="UriLink"/></description></item>
        /// </list>
        /// </remarks>
        public IConfigurationTestBuildViewerLink BuildViewerLink { get; private set; }

        /// <summary>
        /// Gets the configuration for common link.
        /// </summary>
        /// <remarks>
        /// Not null value returns either no of these properties or only one property:
        /// <list type="bullet">
        /// <item><description><see cref="LinkedTemplate"/></description></item>
        /// <item><description><see cref="WorkItemEditorLink"/></description></item>
        /// <item><description><see cref="WorkItemViewerLink"/></description></item>
        /// <item><description><see cref="BuildViewerLink"/></description></item>
        /// <item><description><see cref="UriLink"/></description></item>
        /// </list>
        /// </remarks>
        public IConfigurationTestUriLink UriLink { get; private set; }

        /// <summary>
        /// Gets the configuration for attachment links
        /// </summary>
        /// <remarks>
        /// Not null value returns either no of these properties or only one property:
        /// <list type="bullet">
        /// <item><description><see cref="LinkedTemplate" /></description></item>
        /// <item><description><see cref="WorkItemEditorLink" /></description></item>
        /// <item><description><see cref="WorkItemViewerLink" /></description></item>
        /// <item><description><see cref="BuildViewerLink" /></description></item>
        /// <item><description><see cref="UriLink" /></description></item>
        /// <item><description><see cref="AttachmentLink" /></description></item>
        /// </list>
        /// </remarks>
        public IConfigurationTestAttachmentLink AttachmentLink { get; private set; }

        /// <summary>
        /// Insert a custom WordBookmark on the text of the replacement
        /// </summary>
        public string WordBookmark
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets a value inidcating whether format of ResolutionState should be string
        /// </summary>
        public bool ResolveResolutionState
        {
            get;
            private set;
        }

        #endregion Implementation of IConfigurationTestTemplate
    }
}
