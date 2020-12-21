#region Usings
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.InfoStorage;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Converter
{
    /// <summary>
    /// Converts a UserInformation to a visibility state.
    /// </summary>
    public class UserInformationToTfsButtonVisibilityConverter : IValueConverter
    {
        /// <summary>
        /// Returns collapsed for anything but <see cref="UserInformationType.Conflicting"/>
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return ((IUserInformation)value).Type == UserInformationType.Conflicting
                       ? Enum.Parse(typeof(Visibility), (string)parameter)
                       : Visibility.Collapsed;
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
