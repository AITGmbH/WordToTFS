namespace AIT.TFS.SyncService.Common.Helper
{
    #region Usings
    using System;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Microsoft.Office.Interop.Word;
    #endregion

    /// <summary>
    /// Helper class for pasting HTML strings into a word document.
    /// </summary>
    public static class WordHelper
    {
        #region Consts
        public const int CommandFailedCode = -2146824090;                         // 0x800A1066;
        private const int PropertyNotAvailableDocumentNotActiveCode = -2146823683; // 0x800A11FD;
        #endregion

        /// <summary>
        /// Pastes an html string at the specified range of a word document.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="value">The html string value.</param>
        /// <param name="useBold">Makes the content bold.</param>
        public static void PasteSpecial(Range range, string value, bool useBold = false)
        {
            PasteSpecial(range, value, 5, true, useBold);
        }

        /// <summary>
        /// Pastes an html string at the specified range of a word document.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="value">The html string value.</param>
        /// <param name="tries">The maximum number of tries of pasting the code in the document. See SER's comment below.</param>
        /// <param name="useBold">Wraps the complete html content with a b-tag.</param>
        /// <param name="throwException">if set to <c>true</c> an exception is thrown if pasting does not work.</param>
        private static void PasteSpecial(Range range, string value, int tries, bool throwException, bool useBold)
        {
            // This is experimental code. Word sometimes returns an error when pasting as html
            // so I try it several times here. Sorry about the low code quality - SER
            if (range == null)
            {
                throw new ArgumentNullException("range");
            }
            //Bug 20639 - Replace <br> with self-closing tag <br />
            value = value.Replace("<br>", "<br />");
            
            // Copy data to clipboard
            var clipboardIsAvailable = ClipboardHelper.SetHtmlDataToClipboard(ClipboardHelper.AddHtmlClipboardHeader(value, useBold));

            if (clipboardIsAvailable)
            {
                // Fix for #15480 - Refresh is duplicating html tables
                // Range is trimmed because the text property is returning the end of cell marker
                // For some reason, this results in pasting outside the cell.
                // (Maybe range.end = range.start so adding -1 ...)
                //_range.MoveEnd(WdUnits.wdCharacter, 1);
                //_range.Delete();

                var success = false;
                COMException exception = null;

                while (tries > 0 && !success)
                {
                    try
                    {
                        range.PasteSpecial(DataType: WdPasteDataType.wdPasteHTML);
                        success = true;
                    }
                    catch (COMException ce)
                    {
                        Thread.Sleep(50);
                        exception = ce;
                        if (ce.ErrorCode == PropertyNotAvailableDocumentNotActiveCode ||
                            ce.ErrorCode == CommandFailedCode)
                        {
                            tries--;
                        }
                        else
                        {
                            tries = 0;
                        }
                    }
                }

                if (success == false && exception != null && throwException)
                {
                    throw exception;
                }

                //_range.MoveEnd(WdUnits.wdCharacter, -1);
            }
        }
    }
}
