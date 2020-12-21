#region Usings
using System;
using System.Globalization;
using System.Windows.Data;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.View.Controls.Properties;
#endregion

namespace AIT.TFS.SyncService.View.Controls.TestReport
{
    /// <summary>
    /// Converts the value of <see cref="DocumentStructureType"/> to human text.
    /// </summary>
    public class DocumentStructureToTextConverter : IValueConverter
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
            if (!(value is DocumentStructureType))
                return value;
            var enumValue = (DocumentStructureType)value;
            var retValue = "???";
            switch (enumValue)
            {
                case DocumentStructureType.IterationPath:
                    retValue = Resources.DocumentStructureType_IterationPath;
                    break;
                case DocumentStructureType.AreaPath:
                    retValue = Resources.DocumentStructureType_AreaPath;
                    break;
                case DocumentStructureType.TestPlanHierarchy:
                    retValue = Resources.DocumentStructureType_TestPlanHierarchy;
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
