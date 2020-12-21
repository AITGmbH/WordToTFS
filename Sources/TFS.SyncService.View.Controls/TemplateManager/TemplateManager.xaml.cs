using System;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Model.TemplateManagement;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using Microsoft.Office.Interop.Word;

namespace AIT.TFS.SyncService.View.Controls.TemplateManager
{
    /// <summary>
    /// Interaction logic for TemplateManagerControl
    /// </summary>
    public partial class TemplateManagerControl : IHostPaneControl
    {
        #region Fields

        private readonly Model.TemplateManagement.TemplateManager _templateManager;

        #endregion Fields

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateManagerControl"/> class.
        /// </summary>
        /// <param name="documentModel">Associated document model contains word document.</param>
        /// <param name="templateManager">Template manager contains the currently available template bundles.</param>
        public TemplateManagerControl(ISyncServiceDocumentModel documentModel, Model.TemplateManagement.TemplateManager templateManager)
        {
            InitializeComponent();
            if (documentModel == null)
            {
                throw new ArgumentNullException("documentModel");
            }
            if (templateManager == null)
            {
                throw new ArgumentNullException("templateManager");
            }
            AttachedDocument = documentModel.WordDocument as Document;
            _templateManager = templateManager;
            DataContext = new TemplateManagerModel(_templateManager);
        }

        #endregion Constructor

        #region Implementation of IHostPaneControl interface

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
            // We don't care
        }

        #endregion Implementation of IHostPaneControl interface
    }
}