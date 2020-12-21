#region Usings
using System;
using System.Globalization;
using System.Windows.Data;
using AIT.TFS.SyncService.Contracts.Enums;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Converter
{
    /// <summary>
    /// Converts a user information type into an icon.
    /// </summary>
    public class UserInformationToImagePathConverter : IValueConverter
    {
        /// <summary>
        /// Converts a user information type into an icon.
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var image = string.Empty;

            switch ((UserInformationType)value)
            {
                case UserInformationType.Success:
                    image = "successed.png";
                    break;
                case UserInformationType.Warning:
                    image = "warning.png";
                    break;
                case UserInformationType.Error:
                    image = "failed.png";
                    break;
                case UserInformationType.Unmodified:
                    image = "skiped.png";
                    break;
                case UserInformationType.Conflicting:
                    image = "conflict.png";
                    break;
            }
            return $@"{parameter}/{image}";
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