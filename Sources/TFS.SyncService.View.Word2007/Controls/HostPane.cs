using System.Windows;
using System.Windows.Forms;

namespace TFS.SyncService.View.Word2007.Controls
{
    /// <summary>
    /// Generic wrapper for WPF-UserControls to be displayed in a word forms panel
    /// </summary>
    /// <typeparam name="TControl">Type of the hosted control</typeparam>
    public partial class HostPane<TControl> : UserControl where TControl : UIElement
    {
        /// <summary>
        /// Initializes a new instance of the class.
        /// </summary>
        /// <param name="content">Host control hosted in host pane.</param>
        public HostPane(UIElement content)
        {
            InitializeComponent();
            WPFContentHost.Child = content;
        }

        /// <summary>
        /// Gets or sets the hosted control.
        /// </summary>
        public TControl HostedControl
        {
            get
            {
                return WPFContentHost.Child as TControl;
            }
            set
            {
                WPFContentHost.Child = value;
            }
        }
    }
}