namespace AIT.TFS.SyncService.Adapter.TFS2012.Helper
{
    #region Usings
    using Microsoft.TeamFoundation.WorkItemTracking.Client;
    using System.Globalization;
    using System;
    #endregion

    /// <summary>
    /// Translates team foundation error codes into meaningful text
    /// </summary>
    public static class ErrorHelper
    {
        /// <summary>
        /// Looks up an error text to display to the user for a field status.
        /// </summary>
        /// <param name="field">Field to look up text for status</param>
        /// <returns>A localized text explaining the field status</returns>
        public static string GetFieldStatusMessage(Field field)
        {
            if (field == null)
                throw new ArgumentNullException("field");

            return string.Format(CultureInfo.CurrentCulture, GetErrorText(field), field.Name);
        }

        /// <summary>
        /// Gets an appropriate error message for a given field status
        /// </summary>
        private static string GetErrorText(Field field)
            {
            switch (field.Status)
            {
                case FieldStatus.Valid:
                    return Properties.Resources.FieldStatusValidFieldStatusValid;

                case FieldStatus.InvalidEmpty:
                    return Properties.Resources.FieldStatusInvalidEmpty;

                case FieldStatus.InvalidNotEmpty:
                    return Properties.Resources.FieldStatusInvalidNotEmpty;

                case FieldStatus.InvalidFormat:
                    return Properties.Resources.FieldStatusInvalidFormat;

                case FieldStatus.InvalidListValue:
                    return Properties.Resources.FieldStatusInvalidListValue;

                case FieldStatus.InvalidOldValue:
                    return Properties.Resources.FieldStatusInvalidOldValue;

                case FieldStatus.InvalidNotOldValue:
                    return Properties.Resources.FieldStatusInvalidNotOldValue2;
                    //return Properties.Resources.FieldStatusInvalidNotOldValue;

                case FieldStatus.InvalidEmptyOrOldValue:
                    return Properties.Resources.FieldStatusInvalidEmptyOrOldValue;

                case FieldStatus.InvalidNotEmptyOrOldValue:
                    return Properties.Resources.FieldStatusInvalidNotEmptyOrOldValue;

                case FieldStatus.InvalidValueInOtherField:
                    return Properties.Resources.FieldStatusInvalidValueInOtherField;

                case FieldStatus.InvalidValueNotInOtherField:
                    return Properties.Resources.FieldStatusInvalidValueNotInOtherField;

                case FieldStatus.InvalidDate:
                    return Properties.Resources.FieldStatusInvalidDate;

                case FieldStatus.InvalidTooLong:
                    var maxChars = GetMaxCharacters(field);
                    if (maxChars > 0 && field.Value != null)
                    {
                        return string.Format(CultureInfo.CurrentCulture,
                                             Properties.Resources.FieldStatusInvalidTooLongWithDetails, "{0}",
                                             field.Value.ToString().Length, maxChars);
                    }
                    return Properties.Resources.FieldStatusInvalidTooLong;

                case FieldStatus.InvalidType:
                    return Properties.Resources.FieldStatusInvalidType;

                case FieldStatus.InvalidComputedField:
                    return Properties.Resources.FieldStatusInvalidComputedField;

                case FieldStatus.InvalidPath:
                    return Properties.Resources.FieldStatusInvalidPath;

                case FieldStatus.InvalidCharacters:
                    return Properties.Resources.FieldStatusInvalidCharacters;

                default:
                    return Properties.Resources.FieldStatusInvalidUnknown;
            }
        }

        /// <summary>
        /// Returns the maximum allowed character for a field of a given type
        /// </summary>
        private static int GetMaxCharacters(Field field)
        {
            switch (field.FieldDefinition.FieldType)
            {
                case FieldType.PlainText:
                case FieldType.Html:
                case FieldType.History:
                    return 32000;
                case FieldType.String:
                    return 255;
            }
            return 0;
        }
    }
}
