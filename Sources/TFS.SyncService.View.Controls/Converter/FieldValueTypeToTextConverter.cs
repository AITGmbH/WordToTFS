#region Usings
using System;
using System.Globalization;
using System.Windows.Data;
using AIT.TFS.SyncService.Contracts.Enums;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Converter
{
    /// <summary>
    /// Converts a field value enumeration value into human readable text.
    /// </summary>
    public class FieldValueTypeToTextConverter : IValueConverter
    {
        /// <summary>
        /// Converts a field value enumeration value into human readable text.
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            switch ((FieldValueType)value)
            {
                case FieldValueType.PlainText:
                    return "Plain Text";
                case FieldValueType.HTML:
                    return "Html";
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Not implemented.
        /// </summary>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
