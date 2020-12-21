#region Usings
using System;
using System.Windows.Data;
using System.Windows.Controls;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Converter
{
    /// <summary>
    /// Converter that compares two <see cref="AIT.TFS.SyncService.Contracts.Configuration.QueryImportOption">QueryImportOption</see>.
    /// </summary>
    public class WorkItemTypeToComboBoxItemConverter : IMultiValueConverter
    {
        /// <summary>
        /// Convert the selected work item type to a ComboBox item. If the selected work item type is null, the
        /// first value is selected (this should be the only ComboBoxItem in the ComboBox)
        /// </summary>
        /// <param name="values">An array of length 2. The first item is the selected work item type, the second all available ComboBoxItems</param>
        /// <param name="targetType">Not used.</param>
        /// <param name="parameter">Not used.</param>
        /// <param name="culture">Not used.</param>
        /// <returns>The first ComboBoxItem if the selected work item type is null</returns>
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if(values == null) throw new ArgumentNullException("values");

            if (values.Length == 2)
            {
                // If the selected work item type in the model is null, selected the first entry in the ComboBox
                var value = values[0];
                var items = values[1] as CompositeCollection;
                if (value == null && items != null && items.Count > 0)
                {
                    return items[0];
                }
                return value;
            }
            return null;
        }

        /// <summary>
        /// Convert the selection back to a WorkItemType. If the only ComboBoxItem ("All work item types") is selected,
        /// returns null, otherwise it returns the selected element.
        /// </summary>
        /// <param name="value">Selected element</param>
        /// <param name="targetTypes">Not used.</param>
        /// <param name="parameter">Not used.</param>
        /// <param name="culture">Not used.</param>
        /// <returns>Null if "All work item types" is selected, the selected work item type otherwise.</returns>
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is ComboBoxItem)
            {
                return new object[] { null };
            }
            return new [] { value };
        }
    }
}
