using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using Microsoft.Office.Interop.Word;
using System.Windows.Input;

namespace AIT.TFS.SyncService.View.Controls
{
    using Model.WindowModel;

    /// <summary>
    /// Interaction logic for ErrorWindow.
    /// </summary>
    public partial class ErrorWindow : IHostPaneControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ErrorWindow"/> class.
        /// </summary>
        public ErrorWindow(ErrorModel errorModel, Document attachedDocument)
        {
            AttachedDocument = attachedDocument;
            DataContext = errorModel;
            InitializeComponent();
        }

        /// <summary>
        /// The method is called after the visibility of the containing panel has changed.
        /// </summary>
        /// <param name="isVisible"><c>true</c>if the panel containing this view is shown. False otherwise.</param>
        public void VisibilityChanged(bool isVisible)
        {
            // We don't care
        }

        /// <summary>
        /// Gets Word document attached to control.
        /// </summary>
        public Document AttachedDocument { get; private set; }


        private void CanExecuteFind(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = e.Parameter is IUserInformation;
        }

        private void ExecutedFind(object sender, ExecutedRoutedEventArgs e)
        {
            var pubinfo = e.Parameter as IUserInformation;
            if (pubinfo == null) return;

            if (pubinfo.NavigateToSourceAction != null)
            {
                pubinfo.NavigateToSourceAction.Invoke();
            }
        }

        private void UserControl_Unloaded_1(object sender, System.Windows.RoutedEventArgs e)
        {
            var storageService = SyncServiceFactory.GetService<IInfoStorageService>();
            storageService.ClearAll();
        }
    }
}