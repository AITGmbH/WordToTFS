using System;
using System.Collections.Generic;

namespace AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects
{
    using Contracts.Configuration;
    using Contracts.Enums;
    using Contracts.WorkItemObjects;

    /// <summary>
    /// The split field is a proxy for a field, that reads from and writes to different fields.
    /// When using headers, the resulting work item may contain split fields, where the value
    /// is read from the header field, but written to the actual work item field. The field behaves exactly
    /// like the field in that values are written, and all Properties return values from the write field.
    /// </summary>
    public class SplitField : IField
    {
        private readonly IField _headerField;
        private readonly IField _workItemField;
        private readonly IWorkItem _workItem;

        /// <summary>
        /// Creates a new split field.
        /// </summary>
        /// <param name="workItem">The work item containing this field.</param>
        /// <param name="headerField">The field from which to read values.</param>
        /// <param name="workItemField">The field to which to write values.</param>
        public SplitField(IWorkItem workItem, IField headerField, IField workItemField)
        {
            _workItem = workItem;
            _headerField = headerField;
            _workItemField = workItemField;
        }

        /// <summary>
        /// Gets the reference name of the field.
        /// </summary>
        public string ReferenceName
        {
            get { return _workItemField.ReferenceName; }
        }

        /// <summary>
        /// Gets the name of the field.
        /// </summary>
        public string Name
        {
            get
            {
                throw new NotSupportedException();
            }
        }

        /// <summary>
        /// Gets the value of the read or write field or sets the value of the write field. If the write field is not
        /// null or empty, it value overrides the value of the readfield and the split field does not use the
        /// read field at all.
        /// </summary>
        public string Value
        {
            get { return ActualReadField.Value; }
            set { _workItemField.Value = value; }
        }

        /// <summary>
        /// Gets or sets the value of the field without converting it to string.
        /// </summary>
        public object OriginalValue
        {
            get
            {
                throw new NotImplementedException("Property is only supported for values coming from TFS/VSTS.");
            }
        }

        /// <summary>
        /// Gets or sets a hyperlink. This is used in word only to create links to web access to TFS.
        /// </summary>
        public string Hyperlink
        {
            get { return _workItemField.Hyperlink; }
            set { _workItemField.Hyperlink = value; }
        }

        /// <summary>
        /// Select field.
        /// </summary>
        public void NavigateTo()
        {
            ActualReadField.NavigateTo();
        }

        /// <summary>
        /// Gets the micro document of the read or write field or sets the value of the write field.
        /// If the write field is not null, its value overrides the value of the read field and the
        /// split field does not use the read field at all.
        /// </summary>
        public object MicroDocument
        {
            get { return ActualReadField.MicroDocument; }
            set { _workItemField.MicroDocument = value; }
        }

        /// <summary>
        /// Compares the value of this field to the value of another field.
        /// </summary>
        /// <param name="value">The value to compare to.</param>
        /// <param name="ignoreFormatting">Sets whether formatting is ignored when comparing html fields</param>
        /// <returns>True if the values are equal, False if not.</returns>
        public bool CompareValue(string value, bool ignoreFormatting)
        {
            return ActualReadField.CompareValue(value, ignoreFormatting);
        }

        /// <summary>
        /// Gets whether this field is editable.
        /// </summary>
        public bool IsEditable
        {
            get { return _workItemField.IsEditable; }
        }

        /// <summary>
        /// Gets the configuration of the field from which values are read. For a header field defined with
        /// SetInNewTFSWorkItem this is the header field for new work items, the work item field otherwise.
        /// </summary>
        public IConfigurationFieldItem Configuration
        {
            get { return ActualReadField.Configuration; }
        }

        /// <summary>
        /// Gets or sets a list of allowed values for this field. The values are not explicitly checked but used to populate a word dropdown list.
        /// </summary>
        public IList<string> AllowedValues
        {
            get
            {
                return new List<string>();
            }
            set
            {
                // The allowed values from header fields are invariant because the allowed values of different affected work items might be different (depending on state etc)
            }
        }

        /// <summary>
        /// Refreshes the Bookmarks of the field, this should be used after setting the values
        /// </summary>
        public void RefreshBookmarks()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Returns the field actually used to read from. The write field is
        /// used instead of the read field if it is set and the direction of the
        /// header field is SetInNewTFS
        /// </summary>
        private IField ActualReadField
        {
            get
            {
                if (_headerField.Configuration.Direction == Direction.SetInNewTfsWorkItem && ((!string.IsNullOrEmpty(_workItemField.Value) && _workItem.IsNew) || !_workItem.IsNew)) return _workItemField;
                return _headerField;
            }
        }

        /// <summary>
        /// Gets flag saying whether ole object exists
        /// </summary>
        public bool ContainsOleObject
        {
            get
            {
                throw new NotSupportedException();
            }
        }

        /// <summary>
        /// Gets a value indicating whether this field contains inline shapes like pictures and smart arts.
        /// </summary>
        public bool ContainsShapes
        {
            get
            {
                throw new NotSupportedException();
            }
        }
    }
}