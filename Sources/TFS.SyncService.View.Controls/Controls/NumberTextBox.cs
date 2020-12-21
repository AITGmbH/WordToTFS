#region Usings
using System.Windows.Controls;
using System.Windows.Input;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Controls
{
    /// <summary>
    /// Control implements text box for number input
    /// </summary>
    public class NumberTextBox : TextBox
    {
        #region Private consts

        private const string ConstDefaultNumberValue = "0";

        #endregion Private consts

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="NumberTextBox()"/> class.
        /// </summary>
        public NumberTextBox()
        {
            PreviewKeyDown += HandlePreviewKeyEvent;
            PreviewKeyUp += HandlePreviewKeyEvent;
            LostFocus += HandleLostFocusEvent;
        }

        #endregion Constructors

        #region Private event handlers

        /// <summary>
        /// Occurs when this element loses logical focus.
        /// </summary>
        /// <param name="sender">Sender of the event.</param>
        /// <param name="e">Event data.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Controls.TextBox.set_Text(System.String)", Justification = "This text is number - 0")]
        private void HandleLostFocusEvent(object sender, System.Windows.RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Text))
            {
                Text = ConstDefaultNumberValue;
                return;
            }
            int number;
            if (!int.TryParse(Text, out number))
            {
                Text = ConstDefaultNumberValue;
                return;
            }
        }

        /// <summary>
        /// Occurs when a key is pressed / released while focus is on this element.
        /// </summary>
        /// <param name="sender">Sender of the event.</param>
        /// <param name="e">Event data.</param>
        private void HandlePreviewKeyEvent(object sender, KeyEventArgs e)
        {
            if ((Keyboard.GetKeyStates(Key.LeftCtrl) & KeyStates.Down) == KeyStates.Down)
                return;
            if ((Keyboard.GetKeyStates(Key.RightCtrl) & KeyStates.Down) == KeyStates.Down)
                return;
            if ((Keyboard.GetKeyStates(Key.LeftAlt) & KeyStates.Down) == KeyStates.Down)
                return;
            if ((Keyboard.GetKeyStates(Key.RightAlt) & KeyStates.Down) == KeyStates.Down)
                return;
            if (e == null)
                return;
            e.Handled = true;
            if (e.Key >= Key.D0 && e.Key <= Key.D9)
                e.Handled = false;
            if (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9)
                e.Handled = false;
            if (e.Key >= Key.F1 && e.Key <= Key.F24)
                e.Handled = false;
            switch (e.Key)
            {
                case Key.Back:
                case Key.Delete:
                case Key.Tab:
                case Key.PageUp:
                case Key.PageDown:
                case Key.End:
                case Key.Home:
                case Key.Left:
                case Key.Up:
                case Key.Right:
                case Key.Down:
                case Key.Select:
                    e.Handled = false;
                    break;
            }
        }

        #endregion Private event handlers
    }
}
