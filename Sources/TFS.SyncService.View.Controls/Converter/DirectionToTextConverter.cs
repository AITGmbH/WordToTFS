#region Usings
using System;
using System.Globalization;
using System.Windows.Data;
using AIT.TFS.SyncService.Contracts.Enums;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Converter
{
    /// <summary>
    /// Converts a direction enumeration value into human readable text. 
    /// </summary>
    public class DirectionToTextConverter : IValueConverter
    {
        /// <summary>
        /// Converts a direction into human readable text
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            switch ((Direction)value)
            {
                case Direction.OtherToTfs:
                    return "Other -> Tfs";
                case Direction.SetInNewTfsWorkItem:
                    return "New Workitem";
                case Direction.TfsToOther:
                    return "Tfs -> Other";
                case Direction.GetOnly:
                    return "Get only";
                case Direction.PublishOnly:
                    return "Publish only";
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
