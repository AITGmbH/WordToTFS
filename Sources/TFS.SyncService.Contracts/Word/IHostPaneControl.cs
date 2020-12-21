using Microsoft.Office.Interop.Word;
namespace AIT.TFS.SyncService.View.Controls.Interfaces
{
    /// <summary>
    /// Interface for control which is hosted in HostPane.
    /// </summary>
    public interface IHostPaneControl
    {
        /// <summary>
        /// Gets Word document attached to control.
        /// </summary>
        Document AttachedDocument { get; }

        /// <summary>
        /// The method is called after the visibility of the containing panel has changed.
        /// </summary>
        /// <param name="isVisible"><c>true</c>if the panel containing this view is shown. False otherwise.</param>
        void VisibilityChanged(bool isVisible);
    }
}
