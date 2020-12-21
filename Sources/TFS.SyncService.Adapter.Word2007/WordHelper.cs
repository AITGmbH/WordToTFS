using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Factory;
using Microsoft.Office.Interop.Word;
using System.Globalization;
using System.Runtime.InteropServices;

namespace AIT.TFS.SyncService.Adapter.Word2007
{
    /// <summary>
    /// Helper functions for word documents.
    /// </summary>
    public static class WordSyncHelper
    {
        /// <summary>
        /// Method gets the information if the cursor is in any table positioned.
        /// </summary>
        /// <returns>True if the cursor positioned in a table. Otherwise false.</returns>
        public static bool IsCursorInTable(Selection selection)
        {
            Guard.ThrowOnArgumentNull(selection, "selection");

            try
            {
                if (selection.Tables.Count != 0 && selection.Cells.Count != 0)
                        return true;
            }
            catch (COMException exception)
            {
                SyncServiceTrace.LogException(exception);
            }

            return false;
        }

        /// <summary>
        /// Method gets the information if the cursor is behind any table positioned.
        /// </summary>
        /// <returns>True if the cursor positioned immediately after table. Otherwise false.</returns>
        public static bool IsCursorBehindTable(Selection selection)
        {
            Guard.ThrowOnArgumentNull(selection, "selection");

            try
            {
                if (selection == null)
                {
                    return false;
                }
                if (selection.Tables.Count != 0 && selection.Cells.Count == 0)
                {
                    return true;
                }
            }
            catch (COMException exception)
            {
                SyncServiceTrace.LogException(exception);
            }

            return false;
        }

        /// <summary>
        /// Removes trailing empty or control characters from a string
        /// </summary>
        /// <param name="cellContent">Dirty <see cref="string"/></param>
        /// <returns>Clean <see cref="string"/></returns>
        public static string FormatCellContent(string cellContent)
        {
            int lastLength;
            do
            {
                lastLength = cellContent.Length;
                cellContent = cellContent.Trim();
                cellContent = cellContent.Trim('\a');
            } while (lastLength != cellContent.Length);

            return cellContent;
        }

        /// <summary>
        /// Determines whether range contains OLE object.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns>
        ///   <c>true</c> if range contains OLE object; otherwise, <c>false</c>.
        /// </returns>
        public static bool ContainsRangeOleObject(Range range)
        {
            Guard.ThrowOnArgumentNull(range, "range");

            // check in shapes collection
            if (range.ShapeRange.Count > 0)
            {
                return true;
            }

            // check in functions collection
            if (range.OMaths.Count > 0)
            {
                return true;
            }

            // check in inline shapes collection
            foreach (InlineShape inlineShape in range.InlineShapes)
            {
                // evaluate shape as OLE only if it is not picture
                if (inlineShape.Type != WdInlineShapeType.wdInlineShapePicture &&
                    inlineShape.Type != WdInlineShapeType.wdInlineShapeLinkedPicture)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Gets the range object of the cell at the given position
        /// </summary>
        /// <param name="table">Table from which to get a value.</param>
        /// <param name="row">Row of the cell.</param>
        /// <param name="column">Column of the cell.</param>
        /// <exception cref="ConfigurationException">If there is no cell with the given row and column</exception>
        public static Range GetCellRange(Table table, int row, int column)
        {
            Guard.ThrowOnArgumentNull(table, "table");

            try
            {
                var cell = table.Cell(row, column);
                if (cell != null)
                {
                    return cell.Range;
                }
            }
            catch (COMException comException)
            {
                // Broken configuration.
                throw new ConfigurationException($"The referenced cell (Row={row}, Column={column}) does not exist. Cell referencing does not support merged cells. Make sure the table does not contain merged cells.",comException);
            }

            return null;
        }

        /// <summary>
        /// Extension method to check if a bookmark for a work item exists
        /// </summary>
        public static bool Exists(this Bookmarks bookmarks, int workItemId)
        {
            Guard.ThrowOnArgumentNull(bookmarks, "bookmarks");
            return bookmarks.Exists("w2t" + workItemId);
        }

        /// <summary>
        /// Extension method to lookup bookmark for work item by work item id
        /// </summary>
        public static Bookmark Item(this Bookmarks bookmarks, int workItemId)
        {
            Guard.ThrowOnArgumentNull(bookmarks, "bookmarks");
            return bookmarks["w2t" + workItemId];
        }
    }
}
