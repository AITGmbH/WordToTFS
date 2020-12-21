using AIT.TFS.SyncService.Model.TestReport;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using Microsoft.Office.Interop.Word;

namespace AIT.TFS.SyncService.View.Controls.TestReport
{
    /// <summary>
    /// Interaction logic for TestSpecificationReport.
    /// </summary>
    public partial class TestSpecificationReport : IHostPaneControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TestSpecificationReport"/> class.
        /// </summary>
        public TestSpecificationReport(Document document)
        {
            InitializeComponent();
            AttachedDocument = document;
        }

        /// <summary>
        /// Gets or sets the model of the view.
        /// </summary>
        public TestSpecificationReportModel Model
        {
            get { return DataContext as TestSpecificationReportModel; }
            set { DataContext = value; }
        }

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
    }
}
