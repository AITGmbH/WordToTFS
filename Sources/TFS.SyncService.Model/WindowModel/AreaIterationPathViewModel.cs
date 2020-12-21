#region Usings
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Model.WindowModelBase;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
#endregion

namespace AIT.TFS.SyncService.Model.WindowModel
{
    /// <summary>
    /// ViewModel class for AreaIterationPathView view.
    /// </summary>
    public class AreaIterationPathViewModel : ExtBaseModel
    {
        #region Private fields
        private IAreaIterationNode _selectedAreaIterationNode;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="AreaIterationPathViewModel"/> class.
        /// </summary>
        /// <param name="tfsService">Team foundation server service.</param>
        /// <param name="model">Associated model of word document.</param>
        public AreaIterationPathViewModel(ITfsService tfsService, ISyncServiceDocumentModel model)
        {
            TfsService = tfsService;
            DocumentModel = model;
        }
        #endregion

        #region Public properties
        /// <summary>
        /// Gets or sets team foundation server service.
        /// </summary>
        public ITfsService TfsService
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets associated word document.
        /// </summary>
        public Document WordDocument
        {
            get
            {
                return DocumentModel.WordDocument as Document;
            }
        }

        /// <summary>
        /// Gets model of associated word document.
        /// </summary>
        public ISyncServiceDocumentModel DocumentModel
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets Area and Iteration hierarchy.
        /// </summary>
        public IList<IAreaIterationNode> AreaAndIterationHierarchy
        {
            get
            {
                return new List<IAreaIterationNode>
                                                 {
                                                     TfsService.GetAreaPathHierarchy(),
                                                     TfsService.GetIterationPathHierarchy()
                                                 }; ;
            }
        }

        /// <summary>
        /// Gets or sets actual selected Area/Iteration node in tree view.
        /// </summary>
        public IAreaIterationNode SelectedAreaIterationNode
        {
            get
            {
                return _selectedAreaIterationNode;
            }
            set
            {
                _selectedAreaIterationNode = value;
                OnPropertyChanged(nameof(SelectedAreaIterationNode));
                OnPropertyChanged(nameof(CanInsert));
            }
        }

        /// <summary>
        /// Gets whether is possible to make insert of Area/Iteration path into document.
        /// </summary>
        public bool CanInsert
        {
            get
            {
                return SelectedAreaIterationNode != null && SelectedAreaIterationNode.RootNode != null;
            }
        }

        /// <summary>
        /// Gets or sets persisted Area/Iteration path.
        /// </summary>
        public string PersistedAreaIterationPath
        {
            get
            {
                return DocumentModel.AreaIterationPath;
            }
            set
            {
                DocumentModel.AreaIterationPath = value;
            }
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Perform Area or Iteration path insert into document.
        /// </summary>
        public void DoInsert()
        {
            // persists area iteration path
            PersistedAreaIterationPath = SelectedAreaIterationNode.Path;
            // insert area iteration path into document
            WordDocument.ActiveWindow.Selection.InsertAfter(SelectedAreaIterationNode.RelativePath);
        }
        #endregion
    }
}