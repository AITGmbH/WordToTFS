using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Factory;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// Class that represents a single work item field.
    /// </summary>
    public class ConfigurationFieldItem : IConfigurationFieldItem
    {
        private string _testCaseStepDelimiter = "->";

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationFieldItem"/> class.
        /// </summary>
        /// <param name="referenceName">The field name of the destination field.</param>
        /// <param name="mappingName">The field name of the source field.</param>
        /// <param name="valueType">Determines whether the exported value should be plain text or html.</param>
        /// <param name="direction">Determines the copy direction of the field.</param>
        /// <param name="rowIndex">Table row index.</param>
        /// <param name="colIndex">Table column index.</param>
        /// <param name="testCaseStepDelimiter">Delimiter string for Test Case step between Title and Expected Result.</param>
        /// <param name="handleAsDocument">Flag whether field value is handled as document.</param>
        /// <param name="handleAsDocumentMode">Defines type for HandleAsDocument processing.</param>
        /// <param name="defaultValue">Default value of this field. Null means that no default value defined and don't configure this default value in GUI.</param>
        /// <param name="shapeOnlyWorkaroundMode">The mode used to work around the shape only restriction.</param>
        /// <param name="wordBookmark">a word Bookmark that is inserted at the corresponding field</param>
        public ConfigurationFieldItem(
            string referenceName, string mappingName, FieldValueType valueType, Direction direction, int rowIndex,
            int colIndex, string testCaseStepDelimiter, bool handleAsDocument, HandleAsDocumentType handleAsDocumentMode, string oLEMarkerField, string oLEMarkerValue, IConfigurationFieldItemDefaultValue defaultValue, ShapeOnlyWorkaroundMode shapeOnlyWorkaroundMode, string wordBookmark, string variableName, string dateTimeFormat)
            : this(referenceName, mappingName, valueType, direction, rowIndex, colIndex, testCaseStepDelimiter, handleAsDocument, handleAsDocumentMode, oLEMarkerField, oLEMarkerValue, defaultValue, shapeOnlyWorkaroundMode, true, wordBookmark, variableName, dateTimeFormat)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationFieldItem"/> class.
        /// </summary>
        /// <param name="referenceName">The field name of the destination field.</param>
        /// <param name="mappingName">The field name of the source field.</param>
        /// <param name="valueType">Determines whether the exported value should be plain text or html.</param>
        /// <param name="direction">Determines the copy direction of the field.</param>
        /// <param name="rowIndex">Table row index.</param>
        /// <param name="colIndex">Table column index.</param>
        /// <param name="testCaseStepDelimiter">Delimiter string for Test Case step between Title and Expected Result.</param>
        /// <param name="handleAsDocument">Flag whether field value is handled as document.</param>
        /// <param name="handleAsDocumentMode">Defines type for HandleAsDocument processing.</param>
        /// <param name="defaultValue">Default value of this field. Null means that no default value defined and don't configure this default value in GUI.</param>
        /// <param name="shapeOnlyWorkaroundMode">The mode used to work around the shape only restriction.</param>
        /// <param name="isMapped">Sets whether this field is actually mapped to word</param>
        /// <param name="wordBookmark"></param>
        public ConfigurationFieldItem(string referenceName, string mappingName, FieldValueType valueType, Direction direction, int rowIndex, int colIndex, string testCaseStepDelimiter, bool handleAsDocument, HandleAsDocumentType handleAsDocumentMode, string oLEMarkerField, string oLEMarkerValue, IConfigurationFieldItemDefaultValue defaultValue, ShapeOnlyWorkaroundMode shapeOnlyWorkaroundMode, bool isMapped, string wordBookmark, string variableName, string dateTimeFormat)
        {
            Guard.ThrowOnArgumentNull(referenceName, "referenceName");

            ReferenceFieldName = referenceName;
            SourceMappingName = mappingName;
            FieldValueType = valueType;
            Direction = direction;
            ColIndex = colIndex;
            RowIndex = rowIndex;
            DefaultValue = defaultValue;
            ShapeOnlyWorkaroundMode = shapeOnlyWorkaroundMode;
            DateTimeFormat = dateTimeFormat;
            if (string.IsNullOrEmpty(testCaseStepDelimiter) == false)
            {
                TestCaseStepDelimiter = testCaseStepDelimiter;
            }
            HandleAsDocument = handleAsDocument;
            HandleAsDocumentMode = handleAsDocumentMode;
            OLEMarkerField = oLEMarkerField;
            OLEMarkerValue = oLEMarkerValue;
            IsMapped = isMapped;
            WordBookmark = wordBookmark;
            VariableName = variableName;
        }

        /// <summary>
        /// The field name of the destination field.
        /// </summary>
        public string ReferenceFieldName { get; private set; }

        /// <summary>
        /// The field name of the source field.
        /// </summary>
        public string SourceMappingName { get; set; }

        /// <summary>
        /// Determines whether the exported value should be plain text or html.
        /// </summary>
        public FieldValueType FieldValueType { get; private set; }

        /// <summary>
        /// Determines the format for the date tiem that should be used.
        /// </summary>
        /// <value>
        /// The date time format (for example "dd.MM.yyyy).
        /// </value>
        public string DateTimeFormat { get; private set; }

        /// <summary>
        /// Determines the copy direction of the field.
        /// </summary>
        public Direction Direction { get; private set; }

        /// <summary>
        /// Gets the table column index.
        /// </summary>
        public int ColIndex { get; private set; }

        /// <summary>
        /// Gets the table column index.
        /// </summary>
        public int RowIndex { get; private set; }

        /// <summary>
        /// Gets the delimiter between Title and Expected Result in Test Case step.
        /// </summary>
        public string TestCaseStepDelimiter
        {
            get { return _testCaseStepDelimiter; }
            private set { _testCaseStepDelimiter = value; }
        }

        /// <summary>
        /// Gets flag whether current field value is handled as document.
        /// </summary>
        public bool HandleAsDocument { get; private set; }

        /// <summary>
        /// Gets the handle as document mode.
        /// </summary>
        public HandleAsDocumentType HandleAsDocumentMode { get; private set; }

        /// <summary>
        /// Gets the reference name of field where OLE marker is setted
        /// </summary>
        public string OLEMarkerField { get; private set; }

        /// <summary>
        /// Gets the value of field which used to mark OLE object exists 
        /// </summary>
        public string OLEMarkerValue { get; private set; }

        /// <summary>
        /// Gets the default value object predefined for the field. Null means that no default 
        /// value defined and don't configure this default value in GUI.
        /// </summary>
        public IConfigurationFieldItemDefaultValue DefaultValue { get; private set; }

        /// <summary>
        /// Returns if this field is actually mapped to word
        /// </summary>
        public bool IsMapped { get; private set; }

        /// 
        ///
        ///
        public string VariableName{get;private set;}

        /// <summary>
        /// Returns a clone of this field configuration.
        /// </summary>
        /// <returns></returns>
        public IConfigurationFieldItem Clone()
        {
            // TODO this does not clone default value
            return new ConfigurationFieldItem(ReferenceFieldName, SourceMappingName, FieldValueType, Direction, RowIndex, ColIndex, TestCaseStepDelimiter, HandleAsDocument, HandleAsDocumentMode, OLEMarkerField, OLEMarkerValue, DefaultValue, ShapeOnlyWorkaroundMode, IsMapped,WordBookmark, VariableName, DateTimeFormat);
        }

        /// <summary>
        /// Gets the shape only workaround mode.
        /// </summary>
        public ShapeOnlyWorkaroundMode ShapeOnlyWorkaroundMode
        {
            get;
            private set;
        }


        /// <summary>
        /// Gets a word bookmark
        /// </summary>
        public string WordBookmark
        {
            get;
            private set;
        }


        

    }
}
