using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.WindowModel;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for SettingsWindow.xaml
    /// </summary>
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            DataContext = new SettingsModel();

            InitializeComponent();
            ((SettingsModel)DataContext).DebugLogMessageNeeded += HandleDebugLogMessageNeeded;
            ((SettingsModel)DataContext).ConsoleExtensionIsActivated += HandleActivationOfConsoleExtension;
        }

        /// <summary>
        /// Handle the closing of the dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HandleButtonCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Called when debug log is activated.
        /// </summary>
        private void HandleDebugLogMessageNeeded(object sender, EventArgs e)
        {
            MessageBox.Show(Properties.Resources.DebugLogInfo, "Warning!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        }

        /// <summary>
        /// This handler is called when the checkbox for the installation is activated. It will raise an modal dialog that warns the user that his path is about to get changed. 
        /// </summary>
        private void HandleActivationOfConsoleExtension(object sender, EventArgs e)
        {
                var result = MessageBox.Show(Properties.Resources.ConsoleExtensionInstallationInfo, "Warning!", MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                if (result == MessageBoxResult.Yes)
                {

                   ((SettingsModel)DataContext).InitializeEnvironmentVariables();
                }
        }

        /// <summary>
        /// Open the debug log
        /// </summary>
        private void HandleButtonOpenDebugLogClick(object sender, RoutedEventArgs e)
        {
            var process = Process.Start(SyncServiceTrace.LogFile);
            if (process != null)
            {
                process.Dispose();
            }
        }

        /// <summary>
        /// Handle the click on the button that opens the install path of word to tfs
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HandleButtonInstallPathClick(object sender, RoutedEventArgs e)
        {
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            //CodeBase is the location of the ClickOnce deployment files
            var uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string clickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath);

            var process = Process.Start("explorer.exe", clickOnceLocation);
            if (process != null)
            {
                process.Dispose();
            }
        }
    }
}
