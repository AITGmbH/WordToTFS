namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
	using AIT.TFS.SyncService.Contracts.Enums;

	/// <summary>
    /// Interface defines configuration for one replacement in template.
    /// </summary>
    public interface IConfigurationTestReplacement
    {
        /// <summary>
        /// Gets the name of bookmark to replace.
        /// </summary>
        string Bookmark { get; }

		/// <summary>
		/// Gets the variable name if a variable should be used for replacement.
		/// </summary>
		string VariableName { get; }

		/// <summary>
		/// Gets the <see cref="FieldValueType"/> for the replacment.
		/// </summary>
		FieldValueType FieldValueType
		{
			get;
		}

		/// <summary>
		/// Gets the property to evaluate and gets the value for replacement - used as text to show.
		/// </summary>
		string PropertyToEvaluate
        {
            get;
        }

        /// <summary>
        /// Gets optinal parameters for a property
        /// </summary>
        string Parameters
        {
            get;
        }

        /// <summary>
        /// Gets the type of value provided by evaluated of property.
        /// </summary>
        PropertyValueFormat ValueType { get; }

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
        /// <item><description><see cref="AttachmentLink"/></description></item>
        /// </list>
        /// </remarks>
        string LinkedTemplate { get; }

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
        /// <item><description><see cref="AttachmentLink"/></description></item>
        /// </list>
        /// </remarks>
        IConfigurationTestWorkItemEditorLink WorkItemEditorLink { get; }

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
        /// <item><description><see cref="AttachmentLink"/></description></item>
        /// </list>
        /// </remarks>
        IConfigurationTestWorkItemViewerLink WorkItemViewerLink { get; }

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
        /// <item><description><see cref="AttachmentLink"/></description></item>
        /// </list>
        /// </remarks>
        IConfigurationTestBuildViewerLink BuildViewerLink { get; }

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
        /// <item><description><see cref="AttachmentLink"/></description></item>
        /// </list>
        /// </remarks>
        IConfigurationTestUriLink UriLink { get; }

        /// <summary>
        /// Gets the configuration for attachment links
        /// </summary>
        /// <remarks>
        /// Not null value returns either no of these properties or only one property:
        /// <list type="bullet">
        /// <item><description><see cref="LinkedTemplate"/></description></item>
        /// <item><description><see cref="WorkItemEditorLink"/></description></item>
        /// <item><description><see cref="WorkItemViewerLink"/></description></item>
        /// <item><description><see cref="BuildViewerLink"/></description></item>
        /// <item><description><see cref="UriLink"/></description></item>
        /// <item><description><see cref="AttachmentLink"/></description></item>
        /// </list>
        /// </remarks>
        IConfigurationTestAttachmentLink AttachmentLink { get; }


        /// <summary>
        /// Insert a custom wordbookmark after replacements are finished
        /// </summary>
        string WordBookmark
        {
            get;
        }

        /// <summary>
        /// Gets a value inidcating whether format of ResolutionState should be string
        /// </summary>
        bool ResolveResolutionState
        {
            get;
        }
    }
}
