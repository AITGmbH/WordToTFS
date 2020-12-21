namespace AIT.TFS.SyncService.View.Controls.TemplateManager.Converter
{
    #region Usings
    using System;
    using System.Globalization;
    using System.Windows;
    using System.Windows.Data;
    using AIT.TFS.SyncService.Contracts.TemplateManager;
    #endregion

    /// <summary>
    /// The class implements converter from template state to visibility.
    /// </summary>
    public class TemplateStateVisibilityConverter : IValueConverter
    {
        /// <summary>
        /// Gets or sets a value indicating whether the logic should be inverted.
        /// </summary>
        public bool Invert
        {
            get;
            set;
        }

        #region Implementation of IConverter
        /// <summary>
        /// Converts a value.
        /// </summary>
        /// <param name="value">The value produced by the binding source.</param>
        /// <param name="targetType">The type of the binding target property.</param>
        /// <param name="parameter">The converter parameter to use.</param>
        /// <param name="culture">The culture to use in the converter.</param>
        /// <returns>A converted value. If the method returns null, the valid null value is used.</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var state = (TemplateState)value;
            if (state == TemplateState.ErrorInTemplate || state == TemplateState.NotInitialized)
                return Invert ? Visibility.Visible : Visibility.Hidden;
            return Invert ? Visibility.Hidden : Visibility.Visible;
        }

        /// <summary>
        /// Converts a value.
        /// </summary>
        /// <param name="value">The value that is produced by the binding target.</param>
        /// <param name="targetType">The type to convert to.</param>
        /// <param name="parameter">The converter parameter to use.</param>
        /// <param name="culture">The culture to use in the converter.</param>
        /// <returns>A converted value. If the method returns null, the valid null value is used.</returns>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion Implementation of IConverter
    }
}
