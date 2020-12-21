#region Usings
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using AIT.TFS.SyncService.Contracts.Enums;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Converter
{
    /// <summary>
    /// Converter that takes a SynchronizationState and returns an tool tip
    /// </summary>
    public sealed class SynchronizationStateToToolTipConverter : IValueConverter
    {
        /// <summary>
        /// Converts Synchronization state to tool tip
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            switch ((SynchronizationState)value)
            {
                case SynchronizationState.Unknown:
                    return Properties.Resources.SyncStateToolTip_Loading;

                case SynchronizationState.DivergedWithConflicts:
                    return Properties.Resources.SyncStateToolTip_DivergedWithConflicts;

                case SynchronizationState.DivergedWithoutConflicts:
                    return Properties.Resources.SyncStateToolTip_DivergedWithoutConflict;

                case SynchronizationState.New:
                    return Properties.Resources.SyncStateToolTip_New;

                case SynchronizationState.Differing:
                    return Properties.Resources.SyncStateToolTip_Differing;

                case SynchronizationState.NotImported:
                    return Properties.Resources.SyncStateToolTip_NotImported;

                case SynchronizationState.Outdated:
                    return Properties.Resources.SyncStateToolTip_Outdated;

                case SynchronizationState.UpToDate:
                    return Properties.Resources.SyncStateToolTip_UpToDate;

                case SynchronizationState.Dirty:
                    return Properties.Resources.SyncStateToolTip_Dirty;

                case SynchronizationState.Aborted:
                    return Properties.Resources.SyncStateToolTip_Aborted;
            }

            return DependencyProperty.UnsetValue;
        }

        /// <summary>
        /// Not supported.
        /// </summary>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}