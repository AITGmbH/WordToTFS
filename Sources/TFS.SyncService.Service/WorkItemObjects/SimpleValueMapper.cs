using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Service.Properties;

namespace AIT.TFS.SyncService.Service.WorkItemObjects
{
    /// <summary>
    /// Provides functionality to map fields from the word table to assigned values in TFS.
    /// </summary>
    internal class SimpleValueMapper : IConverter
    {
        /// <summary>
        /// It is not possible 'string.Empty' to use as key in dictionary. We use this constant as 'string.Empty'.
        /// </summary>
        private const string EmptyToken = "___x_x_x___";

        private readonly StringDictionary _otherToTfsMapping = new StringDictionary();
        private readonly StringDictionary _tfsToOtherMapping = new StringDictionary();

        /// <summary>
        /// Initializes a new instance of the <see cref="SimpleValueMapper"/> class.
        /// </summary>
        /// <param name="fieldName">The field for which this converter is used</param>
        public SimpleValueMapper(string fieldName)
        {
            if (fieldName == null)
            {
                throw new ArgumentNullException("fieldName");
            }

            OriginalFieldName = fieldName;
            FieldName = fieldName.ToUpperInvariant();
        }

        /// <summary>
        /// Original defined name of field. No uppercase was made.
        /// </summary>
        private string OriginalFieldName { get; set; }

        #region IConverter Members

        /// <summary>
        /// The reference field name for which the converter has to be used.
        /// </summary>
        public string FieldName { get; private set; }

        /// <summary>
        /// Executes a conversion between the source and the destination value. 
        /// The converter allows the manipulation of the values before they are written into the work items.
        /// </summary>
        /// <param name="value">The source value for the conversion.</param>
        /// <param name="direction">Determines the direction of the conversion.</param>
        /// <returns>The converted value.</returns>
        public string Convert(string value, Direction direction)
        {
            string keyValues = string.Empty;

            string valueToConvert = value;
            if (string.IsNullOrEmpty(valueToConvert))
            {
                valueToConvert = EmptyToken;
            }

            // try to convert the value depending on the direction
            if (Direction.OtherToTfs == direction)
            {
                if (_otherToTfsMapping.ContainsKey(valueToConvert))
                {
                    return _otherToTfsMapping[valueToConvert];
                }
                else
                {
                    foreach (string val in _otherToTfsMapping.Values)
                    {
                        if (string.Equals(val, valueToConvert, StringComparison.OrdinalIgnoreCase))
                            return val;
                    }
                }

                if (_otherToTfsMapping.ContainsValue(valueToConvert))
                    return valueToConvert;

                // Value not found - compose all possible values
                foreach (string key in _otherToTfsMapping.Keys)
                {
                    if (keyValues.Length != 0)
                    {
                        keyValues += ", ";
                    }

                    if (key == EmptyToken)
                    {
                        keyValues += "''";
                    }

                    else
                    {
                        keyValues += "'" + key + "'";
                    }
                }
            }
            else if (Direction.TfsToOther == direction)
            {
                if (_tfsToOtherMapping.ContainsKey(valueToConvert))
                    return _tfsToOtherMapping[valueToConvert];

                // Value not found - compose all possible values
                foreach (string key in _tfsToOtherMapping.Keys)
                {
                    if (keyValues.Length != 0)
                        keyValues += ", ";
                    if (key == EmptyToken)
                        keyValues += "''";
                    else
                        keyValues += "'" + key + "'";
                }
            }

            // the field conversion has not be found for the requested value --> throw exception
            string message = string.Format(
                CultureInfo.CurrentUICulture,
                Resources.TextFormating_NoRequiredValueInConverter,
                value,
                OriginalFieldName,
                keyValues);
            throw new KeyNotFoundException(message);
        }

        #endregion

        /// <summary>
        /// Method adds the mapping pair for the conversion.
        /// </summary>
        /// <param name="TFS">TFS field name.</param>
        /// <param name="other">Other field name.</param>
        public void AddMapping(string TFS, string other)
        {
            string key;
            string value;

            key = other;
            value = TFS;
            if (string.IsNullOrEmpty(key))
                key = EmptyToken;
            if (!_otherToTfsMapping.ContainsKey(key))
                _otherToTfsMapping.Add(key, value);

            key = TFS;
            value = other;
            if (string.IsNullOrEmpty(key))
                key = EmptyToken;
            if (!_tfsToOtherMapping.ContainsKey(key))
                _tfsToOtherMapping.Add(key, value);
        }
    }
}