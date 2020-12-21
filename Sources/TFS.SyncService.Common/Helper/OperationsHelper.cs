namespace AIT.TFS.SyncService.Common.Helper
{
    #region Usings
    using System.Collections.Generic;
    using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
    using Microsoft.Office.Interop.Word;
    #endregion

    /// <summary>
    /// This helper will process pre and post operations on the current document
    /// </summary>
    public static class OperationsHelper
    {
        #region Public static methods

        /// <summary>
        /// Process operations on the given document
        /// </summary>
        /// <param name="operations">The list of pre or postoperations that should be performed</param>
        /// <param name="document">The document where the operations should be performed</param>
        public static void ProcessOperations(IList<IConfigurationTestOperation> operations,Document document )
        {
            if (operations == null || operations.Count == 0 || document == null) return;

            foreach (var operation in operations)
            {
                switch (operation.Type)
                {
                    case OperationType.InsertParagraph:
                        document.ActiveWindow.Selection.InsertParagraph();
                        break;
                    case OperationType.MoveCursorToStart:
                        document.ActiveWindow.Selection.GoTo(WdGoToItem.wdGoToLine, WdGoToDirection.wdGoToFirst);
                        break;
                    case OperationType.MoveCursorToEnd:
                        document.ActiveWindow.Selection.GoTo(WdGoToItem.wdGoToLine, WdGoToDirection.wdGoToLast);
                        break;
                    case OperationType.MoveCursorToLeft:
                        document.ActiveWindow.Selection.Move(WdUnits.wdCharacter, -1);
                        break;
                    case OperationType.MoveCursorToRight:
                        document.ActiveWindow.Selection.Move(WdUnits.wdCharacter, -1);
                        break;
                    case OperationType.DeleteCharacterLeft:
                        document.ActiveWindow.Selection.Delete(WdUnits.wdCharacter, -1);
                        break;
                    case OperationType.DeleteCharacterRight:
                        document.ActiveWindow.Selection.Delete(WdUnits.wdCharacter, -1);
                        break;
                    case OperationType.InsertNewPage:
                        document.ActiveWindow.Selection.InsertNewPage();
                        break;
                    case OperationType.RefreshAllFieldsInDocument:
                        UpdateAllFieldsInDocument(document);
                        break;
                }
            }
        }
        #endregion

        #region Private static methods
        /// <summary>
        /// Helper method that will update all fields in the document, including headers and footers
        /// </summary>
        private static void UpdateAllFieldsInDocument(Document document)
        {
            if (document.TablesOfFigures != null)
            {
                for (int i = 1; i <= document.TablesOfFigures.Count; i++)
                {
                    var tableOfFigure = document.TablesOfFigures[i];
                    if (tableOfFigure != null)
                    {
                        tableOfFigure.Update();
                    }
                } 
            }

            if (document.TablesOfContents != null)
            {
                for (int i = 1; i <= document.TablesOfContents.Count; i++)
                {
                    var tableOfContent = document.TablesOfContents[i];
                    if (tableOfContent != null)
                    {
                        tableOfContent.Update();
                    }
                } 
            }

            //Update all fields
            document.Fields.Update();            

             // Loop through all sections and update all headers and footers
            foreach (Section section in document.Sections)
            {
                //Update Header
                var headers = section.Headers;
                foreach (HeaderFooter header in headers)
                {
                    header.Range.Fields.Update();
                }

                //Update Footer
                var footers = section.Footers;
                foreach (HeaderFooter footer in footers)
                {
                    footer.Range.Fields.Update();
                }
            }
        }
        #endregion
    }
}
