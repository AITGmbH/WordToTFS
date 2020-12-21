#region Usings
using AIT.TFS.SyncService.Contracts.Enums;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Converter
{
    /// <summary>
    /// Converts a Synchronization state to an image representation
    /// </summary>
    public sealed class SynchronizationStateToImageConverter : IValueConverter
    {
        /// <summary>
        /// Converts a Synchronization state to an image representation
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            switch ((SynchronizationState) value)
            {
                case SynchronizationState.Unknown:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/Loading.png";

                case SynchronizationState.DivergedWithConflicts:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/DivergedWithConflict.png";

                case SynchronizationState.DivergedWithoutConflicts:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/DivergedWithoutConflict.png";

                case SynchronizationState.New:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/New.png";

                case SynchronizationState.Differing:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/Differing.png";

                case SynchronizationState.NotImported:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/NotImported.png";

                case SynchronizationState.Outdated:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/Outdated.png";

                case SynchronizationState.UpToDate:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/UpToDate.png";

                case SynchronizationState.Dirty:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/Dirty.png";
            
                case SynchronizationState.Aborted:
                    return "pack://application:,,,/TFS.SyncService.View.Controls;component/Resources/Images/SynchronizationStates/Aborted.png";
            }

            return DependencyProperty.UnsetValue;
        }

        /// <summary>
        /// Not supported
        /// </summary>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}