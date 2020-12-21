using System.Windows;
using Window = System.Windows.Window;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for OverwriteDialog.xaml
    /// </summary>
    public partial class OverwriteDialog : Window
    {
        /// <summary>
        /// Main Constructor
        /// </summary>
        public OverwriteDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Handles the Yes Button (Users wants to overwrite the html content with plaintext
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnYesButtonClick(object sender, RoutedEventArgs e)
        {

            DialogResult = true;
            Close();
        }

        /// <summary>
        /// Handles the No Button (User wants the html synced to tfs
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void OnNoButtonClick(object sender, RoutedEventArgs e)
        {

            DialogResult = false;
            Close();
        }

    }
}
