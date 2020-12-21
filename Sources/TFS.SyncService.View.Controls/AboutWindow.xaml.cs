using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Navigation;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Model.WindowModel;
using MessageBox = System.Windows.MessageBox;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Class implements the about box for word plugin.
    /// </summary>
    public partial class AboutWindow : Window
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AboutWindow"/> class.
        /// </summary>
        public AboutWindow()
        {
            DataContext = new AboutModel();
            InitializeComponent();
        }

        /// <summary>
        /// Called when the user clicks the button 'buttonOK'.
        /// </summary>
        private void HandleButtonOkClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Called when navigation events are requested.
        /// </summary>
        private void HandleHyperlinkNavigate(object sender, RequestNavigateEventArgs e)
        {
            string absoluteUri = e.Uri.AbsoluteUri;
            var process = Process.Start(new ProcessStartInfo(absoluteUri));
            if (process != null)
            {
                process.Dispose();
            }

            e.Handled = true;
        }

        private void HandleLogoDoubleClick(object sender, MouseButtonEventArgs mouseButtonEventArgs)
        {
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            //Location is where the assembly is run from
            string assemblyLocation = assemblyInfo.Location;

            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string clickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath);

            string localAppData= Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                    Constants.ApplicationCompany, Constants.ApplicationName);

            MessageBox.Show($"Assembly location is:{assemblyLocation} \n ClickOnceLocation is: {clickOnceLocation} \n LocalAppData is: {localAppData} \n  ");
          }
    }
}