using System;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using Microsoft.Office.Interop.Word;

namespace AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects
{
    public class StaticValueField : IField
    {
        private readonly Range _range;
        private readonly string _variableName ;

        /// <summary>
        /// A static value field that supports the us of variables in the config of WordToTFS
        /// </summary>
        /// <param name="range"></param>
        /// <param name="configurationFieldItem"></param>
        /// <param name="cutLastMarker"></param>
        public StaticValueField(Range range, IConfigurationFieldItem configurationFieldItem, bool cutLastMarker)
        {
            Guard.ThrowOnArgumentNull(configurationFieldItem, "configurationFieldItem");

            Configuration = configurationFieldItem.Clone();
            string variableName = configurationFieldItem.VariableName;
            if (variableName == null)
            {
                throw new ConfigurationException("You are trying to use a static text field without an assigned variable. Please specify a variable in the config");
            }

            _range = range;
            _variableName = variableName;

            // exclude end-of-cell marker
            if (cutLastMarker && _range != null)
            {
               _range.MoveEnd(WdUnits.wdCharacter, -1);
            }

            if (configurationFieldItem.WordBookmark != null && configurationFieldItem.WordBookmark != " ")
            {

                if (_range != null && _range.Text != null && (_range.Text.Equals("") || _range.Text.Equals(" ") || _range.Text == null))
                {
                    //If the Text of the range is empty, add some non blank space to help word with its Bookmarks
                        _range.Text = "&nbsp;";
                }

                if (_range != null && !_range.Bookmarks.Exists(configurationFieldItem.WordBookmark))
                {
                    _range.Bookmarks.Add(configurationFieldItem.WordBookmark);
                }
            }
        }

        /// <summary>
        /// Gets the reference name of the field.
        /// </summary>
        public string ReferenceName
        {
            get
            {
                return Configuration.ReferenceFieldName;
            }
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
        /// Gets or sets the value of the field.
        /// </summary>
        public string Value
        {
            get
            {
                return _range.Text;
            }

            set
            {
                //Set the value
                _range.Text = value;
                RefreshBookmarks();
            }
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
        /// Get the assigned VariableName
        /// </summary>
        public string VariableName
        {
            get
            {
                return _variableName;
            }
        }
        /// <summary>
        /// Gets or sets a hyperlink. This is used in word only to create links to web access to TFS.
        /// </summary>
        public string Hyperlink
        {
            get;
            set;
        }

        /// <summary>
        /// Navigates the user to this field.
        /// </summary>
        public void NavigateTo()
        {
            if (_range != null)
            {
                _range.Select();
            }
        }

        /// <summary>
        /// Gets or sets the micro document that exports the field value.
        /// </summary>
        public object MicroDocument
        {
            get;
            set;
        }

        /// <summary>
        /// Compares the value of this field to the value of another field.
        /// </summary>
        /// <param name="value">The value to compare to.</param>
        /// <param name="ignoreFormatting">Sets whether formatting is ignored when comparing html fields</param>
        /// <returns>True if the values are equal, False if not.</returns>
        public bool CompareValue(string value, bool ignoreFormatting)
        {
            throw new NotSupportedException("The static value field field does not support comparision. Use the tfs field implementation instead.");
        }

        /// <summary>
        /// Gets whether this field is editable.
        /// </summary>
        public bool IsEditable
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Gets the configuration used to create this field and that defines its behavior.
        /// </summary>
        public IConfigurationFieldItem Configuration
        {
            get;
            private set;
        }

        /// <summary>
        ///  Gets or sets a list of allowed values for this field. The values are not explicitly checked but used to populate a word dropdown list.
        /// </summary>
        public IList<string> AllowedValues
        {
            get;
            set;
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

        /// <summary>
        /// Refreshes the Bookmarks of the field, this should be used after setting the values
        /// </summary>
        public void RefreshBookmarks()
        {
            if (Configuration.WordBookmark != null && !Configuration.WordBookmark.Equals(""))
            {
                _range.Bookmarks.Add(Configuration.WordBookmark);
            }
        }
    }
}
