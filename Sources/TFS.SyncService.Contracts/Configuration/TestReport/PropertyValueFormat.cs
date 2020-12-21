namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    using Enums;

    /// <summary>
    /// Enumeration defines the format of the evaluated property.
    /// </summary>
    public enum PropertyValueFormat
    {
        /// <summary>
        /// Value provided by evaluating of property is plain text.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "PlainText", Justification = "This is the official spelling for TFS. We just use the same")]
        PlainText = 0,

        /// <summary>
        /// Value provided by evaluating of property is html formatted text.
        /// </summary>
        /// <remarks>Text <c>HTML</c> used to get the same text as in <see cref="FieldValueType"/> defined.</remarks>
        HTML,

        /// <summary>
        /// Value provided by evaluating of property is html formatted text and adds bold to the complete Value.
        /// </summary>
        /// <remarks>Text <c>HTML</c> used to get the same text as in <see cref="FieldValueType"/> defined.</remarks>
        HTMLBold,
    }
}
