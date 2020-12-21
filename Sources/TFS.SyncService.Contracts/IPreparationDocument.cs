using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIT.TFS.SyncService.Contracts
{
    /// <summary>
    /// Interface defines functionality of preparation document for long-term operations
    /// </summary>
    public interface IPreparationDocument
    {
        /// <summary>
        /// Prepare the document for the long-term operation. 
        /// Disable Spellchecking and Pagination. This is necessary, because Word Add-ins are not supposed to create big documents in a self running manner
        /// </summary>
        void PrepareDocumentForLongTermOperation();

        /// <summary>
        /// Undo the changes to the document 
        /// Enable Spellchecking and Pagination.
        /// </summary>
        void UndoPreparationsDocumentForLongTermOperation();
    }
}
