namespace AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects
{
    using Contracts.Configuration;
    using Contracts.WorkItemObjects;

    /// <summary>
    /// This class represents a single field of a WordTableWorkItem that is not actually mapped to word.
    /// It is used as dummy for when using the direction "SetInNewTFS" and if a header field is not present
    /// in the work item configuration.
    /// When creating a TFS work item, its configuration must match the word work item field definitions.
    /// With the headers, field configurations are dynamic and the actual field setup of a of work item
    /// may depend on its position within the document. Hidden fields are used to save this dynamic field
    /// configurations.
    /// </summary>
    public class HiddenField : WordTableField
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HiddenField"/> class.
        /// </summary>
        /// <param name="configurationFieldItem">Configuration for this field.</param>
        /// <param name="converter">Converter for values of this field.</param>
        public HiddenField(IConfigurationFieldItem configurationFieldItem, IConverter converter)
            : base(null, configurationFieldItem, converter, false)
        {
        }

        /// <summary>
        /// Gets or sets the value of the field.
        /// </summary>
        public override string Value { get; set; }

        /// <summary>
        /// Gets or sets the micro document that exports the field value.
        /// </summary>
        public override object MicroDocument { get; set; }
    }
}