using System.Windows;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for MissingFieldsInformationWindow.
    /// </summary>
    public partial class MissingFieldsInformationWindow : Window
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MissingFieldsInformationWindow"/> class.
        /// </summary>
        public MissingFieldsInformationWindow()
        {
            InitializeComponent();
        }

        private void OnOkButtonClick(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
