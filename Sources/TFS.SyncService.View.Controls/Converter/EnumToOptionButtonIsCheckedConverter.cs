#region Usings
using System;
using System.Windows.Data;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Converter
{
    /// <summary>
    /// Converter that compares two <see cref="AIT.TFS.SyncService.Contracts.Configuration.QueryImportOption">QueryImportOption</see>.
    /// </summary>
    public class EnumToOptionButtonIsCheckedConverter : IValueConverter
    {
        /// <summary>
        /// Return true if the bound property equals the given parameter.
        ///  </summary>
        /// <param name="value">Value of the bound property.</param>
        /// <param name="targetType">Not used.</param>
        /// <param name="parameter">Enumeration value to compare to.</param>
        /// <param name="culture">Not used.</param>
        /// <returns></returns>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if(value == null) throw new ArgumentNullException("value");

            return value.Equals(parameter);
        }

        /// <summary>
        /// Returns the parameter if the value is true (which is the case only for one IsChecked-Property)
        /// </summary>
        /// <param name="value">Value of the IsChecked-Property.</param>
        /// <param name="targetType">Not used.</param>
        /// <param name="parameter">Enumeration value to compare to.</param>
        /// <param name="culture">Not used.</param>
        /// <returns>Value of the parameter.</returns>
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null) throw new ArgumentNullException("value");

            return value.Equals(false) ? Binding.DoNothing : parameter;
        }
    }
}
