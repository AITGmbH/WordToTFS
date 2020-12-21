using System.Windows;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for WronglyMappedFieldsInformationWindow.
    /// </summary>
    public partial class WronglyMappedFieldsInformationWindow : Window
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WronglyMappedFieldsInformationWindow"/> class.
        /// </summary>
        public WronglyMappedFieldsInformationWindow()
        {
            InitializeComponent();
        }

        private void OnOkButtonClick(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
