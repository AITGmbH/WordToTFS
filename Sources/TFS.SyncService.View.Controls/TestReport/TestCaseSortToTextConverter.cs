#region Usings
using System;
using System.Globalization;
using System.Windows.Data;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.Model.TestReport;
#endregion

namespace AIT.TFS.SyncService.View.Controls.TestReport
{
    /// <summary>
    /// Converts the value of <see cref="TestCaseSortType"/> to human text.
    /// </summary>
    public class TestCaseSortToTextConverter : IValueConverter
    {
        #region Implementation of IValueConverter

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
            if (!(value is TestCaseSortType))
                return value;
            var enumValue = (TestCaseSortType)value;
            var retValue = "???";
            switch (enumValue)
            {
                case TestCaseSortType.None:
                    retValue = Properties.Resources.TestCasesSortType_None;
                    break;
                case TestCaseSortType.IterationPath:
                    retValue = Properties.Resources.TestCasesSortType_IterationPath;
                    break;
                case TestCaseSortType.AreaPath:
                    retValue = Properties.Resources.TestCasesSortType_AreaPath;
                    break;
                case TestCaseSortType.WorkItemId:
                    retValue = Properties.Resources.TestCasesSortType_WorkItemId;
                    break;
            }
            return retValue;
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

        #endregion Implementation of IValueConverter
    }
}
