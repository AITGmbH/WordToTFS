using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Interface defines functionality for one field in work item.
    /// </summary>
    public interface IConfigurationFieldItem
    {
        /// <summary>
        /// The field name of the destination field
        /// </summary>
        string ReferenceFieldName { get; }

        /// <summary>
        /// The field name of the source field
        /// </summary>
        string SourceMappingName { get; set; }

        /// <summary>
        /// Determines whether the exported value should be plain text or html
        /// </summary>
        FieldValueType FieldValueType { get; }

        /// <summary>
        /// Determines the format for the date tiem that should be used.
        /// </summary>
        /// <value>
        /// The date time format (for example "dd.MM.yyyy).
        /// </value>
        string DateTimeFormat { get;}

        /// <summary>
        /// Determines the copy direction of the field
        /// </summary>
        Direction Direction { get; }

        /// <summary>
        /// Table column index
        /// </summary>
        int ColIndex { get; }

        /// <summary>
        /// Table column index
        /// </summary>
        int RowIndex { get; }

        /// <summary>
        /// Gets the delimiter between Title and Expected Result in Test Case step.
        /// </summary>
        string TestCaseStepDelimiter { get; }

        /// <summary>
        /// Gets flag whether current field value is handled as document.
        /// </summary>
        bool HandleAsDocument { get; }

        /// <summary>
        /// Gets the handle as document mode.
        /// </summary>
        HandleAsDocumentType HandleAsDocumentMode { get; }

        /// <summary>
        /// Gets the reference name of field where OLE marker is setted
        /// </summary>
        string OLEMarkerField { get; }

        /// <summary>
        /// Gets the value of field which used to mark OLE object exists 
        /// </summary>
        string OLEMarkerValue { get; }

        /// <summary>
        /// Gets the default value object predefined for the field. Null means that no default value defined and don't configure this default value in GUI.
        /// </summary>
        IConfigurationFieldItemDefaultValue DefaultValue { get; }

        /// <summary>
        /// Returns if this field is actually mapped to word
        /// </summary>
        bool IsMapped { get; }

        /// <summary>
        /// The value of the variable name.
        /// </summary>
        string VariableName { get;}

        /// <summary>
        /// Returns a clone of this field configuration.
        /// </summary>
        IConfigurationFieldItem Clone();

        /// <summary>
        /// Gets the shape only workaround mode.
        /// </summary>
        ShapeOnlyWorkaroundMode ShapeOnlyWorkaroundMode
        {
            get;
        }

        /// <summary>
        /// A Word Bookmark that is automaticly inserted into the corresponding cell of the table
        /// </summary>
        string WordBookmark
        {
            get;
        }


    }
}