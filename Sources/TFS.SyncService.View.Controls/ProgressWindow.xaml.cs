using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.WindowModel;
using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for ProgressWindow.
    /// </summary>
    public sealed partial class ProgressWindow : IDisposable
    {
        private static ProgressWindow _window;

        private readonly ProgressModel _model = new ProgressModel();
        /// <summary>
        /// Initializes a new instance of the <see cref="ProgressWindow"/> class.
        /// </summary>
        /// <remarks>
        /// Call this method in GUI thread.
        /// </remarks>
        public ProgressWindow()
        {
            InitializeComponent();
            var service = SyncServiceFactory.GetService<IProgressService>();
            if (service != null)
                service.AttachModel(_model);
            DataContext = _model;
            KeyDown += ProgressWindow_KeyDown;
        }

        #region IDisposable Members

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            var service = SyncServiceFactory.GetService<IProgressService>();
            if (service != null)
                service.DetachModel();
        }

        #endregion

        /// <summary>
        /// Method shows the progress window in background worker.
        /// </summary>
        public static void ShowWindow(IntPtr parentWindowHandle)
        {
            if (_window != null)
                HideWindow();
            _window = new ProgressWindow();
            if (parentWindowHandle != null)
            {
                var wndInteropHelper = new WindowInteropHelper(_window);
                wndInteropHelper.Owner = parentWindowHandle;
            }
            _window.Show();
        }

        /// <summary>
        /// Method hides the progress window.
        /// </summary>
        public static void HideWindow()
        {
            if (_window == null || _window.Dispatcher == null)
                return;
            if (!_window.Dispatcher.CheckAccess())
            {
                _window.Dispatcher.Invoke(new Action(HideWindow));
                return;
            }
            _window.Close();
            _window = null;
        }

        /// <summary>
        /// Called when a button is clicked.
        /// </summary>
        /// <param name="sender">The object where the event handler is attached.</param>
        /// <param name="e">The event data.</param>
        private void HandleButtonCancelClick(object sender, RoutedEventArgs e)
        {
            _model.ProgressCanceled = true;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            _model.ProgressCanceled = true;
            if (_window != null)
                _window.Dispose();
        }

        private void ProgressWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                _model.ProgressCanceled = true;
        }
    }
}