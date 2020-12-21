using AIT.TFS.SyncService.Model.TestReport;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using Microsoft.Office.Interop.Word;

namespace AIT.TFS.SyncService.View.Controls.TestReport
{
    /// <summary>
    /// Interaction logic for TestResultReport.
    /// </summary>
    public partial class TestResultReport : IHostPaneControl
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TestResultReport"/> class.
        /// </summary>
        public TestResultReport(Document document)
        {
            InitializeComponent();
            AttachedDocument = document;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets or sets the model of the view.
        /// </summary>
        public TestResultReportModel Model
        {
            get { return DataContext as TestResultReportModel; }
            set { DataContext = value; }
        }

        #endregion Public properties

        #region Implementation of IHostPaneControl

        /// <summary>
        /// Gets or sets Word Document instance.
        /// </summary>
        public Document AttachedDocument { get; private set; }

        /// <summary>
        /// The method is called after the visibility of the containing panel has changed.
        /// </summary>
        /// <param name="isVisible"><c>true</c>if the panel containing this view is shown. False otherwise.</param>
        public void VisibilityChanged(bool isVisible)
        {
            if (Model != null)
                Model.VisibilityChanged(isVisible);
        }

        #endregion Implementation of IHostPaneControl
    }
}
