// ReSharper disable StringIndexOfIsCultureSpecific.2
// ReSharper disable UnusedAutoPropertyAccessor.Local
// ReSharper disable StringIndexOfIsCultureSpecific.1
// ReSharper disable StringLastIndexOfIsCultureSpecific.1
namespace AIT.TFS.SyncService.Common.Helper
{
    #region Usings
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    using Microsoft.Office.Interop.Word;
    using System.Text.RegularExpressions;
    using System.IO;

    //using ImageComposer;
    #endregion

    /// <summary>
    /// An internal helper class which simplifies the access to the cf_html format as it is saved on the clipboard
    /// Format of this header is described at: http://msdn.microsoft.com/en-us/library/aa767917(VS.85).aspx
    /// </summary>
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "Windows API, use of standard c naming.")]
    [SuppressMessage("Microsoft.StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Windows API, use of standard c naming.")]
    public class ClipboardHelper
    {
        #region Fields
        private const string ConstHtmlFormat = "Html Format";
        private string _html;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="ClipboardHelper"/> class.
        /// </summary>
        /// <param name="range">Range to process.</param>
        public ClipboardHelper(Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException("range");
            }

            var oldData = GetHtmlDataFromClipboard();

            CopyNodeDataToClipboard(range);
            ClipBoardStream = GetHtmlDataFromClipboard();

            SetHtmlDataToClipboard(oldData);
            this.ParseClipBoardStream(range);
        }
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the byte count from the beginning of the clipboard to the start of the context, or -1 if no context.
        /// </summary>
        private int StartHtml { get; set; }

        /// <summary>
        /// Gets or sets the byte count from the beginning of the clipboard to the end of the context, or -1 if no context.
        /// </summary>
        private int EndHtml { get; set; }

        /// <summary>
        /// Gets or sets the whole cf_html compatible string
        /// </summary>
        private string ClipBoardStream { get; set; }

        /// <summary>
        /// Gets the html stream
        /// </summary>
        public string Html
        {
            get
            {
                if (null == _html)
                {
                    // The clip board values could not be extracted for end html --> return empty string
                    _html = string.Empty;
                    if (-1 != StartHtml || -1 != EndHtml)
                    {
                        _html = RemoveClipBoardHeaderFromHtmlString(ClipBoardStream);
                    }
                }

                return _html;
            }
        }

        #endregion

        #region Public static methods
        /// <summary>
        /// Set html stream into clipboard.
        /// </summary>
        /// <param name="htmlText">HTML stream in string.</param>
        [SuppressMessage("Microsoft.StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "c style PInvoke")]
        public static bool SetHtmlDataToClipboard(string htmlText)
        {
            var cfHtml = NativeMethods.RegisterClipboardFormatA(ConstHtmlFormat);
            if (OpenClipboard())
            {
                NativeMethods.EmptyClipboard();
                var htmlBytes = Encoding.UTF8.GetBytes(htmlText);

                var hglbCopy = NativeMethods.GlobalAlloc(0x0002, new UIntPtr((uint)htmlBytes.Length));
                if (hglbCopy != IntPtr.Zero)
                {
                    var lptstrCopy = NativeMethods.GlobalLock(hglbCopy);
                    if (lptstrCopy != IntPtr.Zero)
                    {
                        Marshal.Copy(htmlBytes, 0, lptstrCopy, htmlBytes.Length);
                        NativeMethods.GlobalUnlock(hglbCopy);
                        NativeMethods.SetClipboardData(cfHtml, hglbCopy);
                    }
                }

                NativeMethods.CloseClipboard();
            }

            return NativeMethods.IsClipboardFormatAvailable(cfHtml);
        }
        #endregion

        #region Private static methods
        /// <summary>
        /// This methods extracts the header information of the cf_html and exports them as properties
        /// </summary>
        private void ParseClipBoardStream(Range range)
        {
            if (!string.IsNullOrEmpty(this.ClipBoardStream))
            {
                var parser = new ImageComposer.CopyStreamParser();
                this.ClipBoardStream = parser.ParseAndRepairForCopy(this.ClipBoardStream, range);
                using (var sr = new StringReader(this.ClipBoardStream))
                {
                    if (HasHeader(sr.ReadLine()))
                    {
                        this.ParseStartHtml(sr.ReadLine());
                        this.ParseEndHtml(sr.ReadLine());
                        ParseStartFragment(sr.ReadLine());
                        ParseEndFragment(sr.ReadLine());
                    }
                }
            }
            else
            {
                this.StartHtml = -1;
                this.EndHtml = -1;
            }
        }

        /// <summary>
        /// Checks the header of the given token.
        /// </summary>
        /// <param name="token">Token to check.</param>
        private static bool HasHeader(string token)
        {
            return token.StartsWith("Version", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Check the start of the given fragment.
        /// </summary>
        /// <param name="token">Fragment to check.</param>
        private static void ParseStartFragment(string token)
        {
            if (token == null)
            {
                throw new ArgumentNullException("token");
            }
            if (!token.StartsWith("StartFragment:", StringComparison.OrdinalIgnoreCase))
                throw new FormatException("Expected token: StartFragment");
        }

        /// <summary>
        /// Check the end of the given fragment.
        /// </summary>
        /// <param name="token">Fragment to check.</param>
        private static void ParseEndFragment(string token)
        {
            if (token == null)
            {
                throw new ArgumentNullException("token");
            }
            if (!token.StartsWith("EndFragment:", StringComparison.OrdinalIgnoreCase))
                throw new FormatException("Expected token: EndFragment");
        }
        
        /// <summary>
        /// Copies the data stream of the range to the clipboard
        /// </summary>
        /// <param name="range">The range which holds the html data</param>
        private static void CopyNodeDataToClipboard(Range range)
        {
            if (OpenClipboard())
            {
                NativeMethods.EmptyClipboard();
                NativeMethods.CloseClipboard();
            }

            if (range.Start != range.End)
            {
                range.FormattedText.Copy();
            }
        }

        /// <summary>
        /// Opens the clipboard and returns the current html stream
        /// </summary>
        /// <returns></returns>
        private static string GetHtmlDataFromClipboard()
        {
            var cfHtml = NativeMethods.RegisterClipboardFormatA(ConstHtmlFormat);
            var stream = string.Empty;
            if (NativeMethods.IsClipboardFormatAvailable(cfHtml) && OpenClipboard())
            {
                var hGMem = NativeMethods.GetClipboardData(cfHtml);
                var pMfp = NativeMethods.GlobalLock(hGMem);
                var len = NativeMethods.GlobalSize(hGMem);
                var bytes = new byte[len.ToUInt32()];
                if (pMfp != IntPtr.Zero)
                {
                    Marshal.Copy(pMfp, bytes, 0, (int)len.ToUInt32());
                    stream = Encoding.UTF8.GetString(bytes);
                }

                NativeMethods.GlobalUnlock(hGMem);
                NativeMethods.CloseClipboard();

            }

            return stream;
        }

        /// <summary>
        /// Open clipboard.
        /// </summary>
        /// <returns></returns>
        private static bool OpenClipboard()
        {
            // try to open it more times
            for (var index = 0; index < 5; index++)
            {
                if (NativeMethods.OpenClipboard(NativeMethods.GetConsoleWindow()))
                {
                    return true;
                }

                // if not successful then try to close it for sure
                NativeMethods.CloseClipboard();
                // wait a while
                Thread.Sleep(50);
            }

            return false;
        }

        /// <summary>
        /// Adds Xhtml header information to Html data string so that it can be placed on clipboard
        /// </summary>
        /// <param name="htmlString">
        /// Html string to be placed on clipboard with appropriate header
        /// </param>
        /// <param name="useBold">Can be used to make the html content bold.</param>
        /// <returns>
        /// String wrapping htmlString with appropriate Html header
        /// </returns>
        internal static string AddHtmlClipboardHeader(string htmlString, bool useBold)
        {
            if (useBold) htmlString = "<b>" + htmlString + "</b>";
            if (htmlString.Contains("Version:1.0\nStartHTML:")) return htmlString;

            var stringBuilder = new StringBuilder();

            // each of 6 numbers is represented by "{0:D10}" in the format string
            // must actually occupy 10 digit positions ("0123456789")
            const string HtmlHeader = "Version:1.0\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\nStartSelection:{4:D10}\r\nEndSelection:{5:D10}\r\n";
            const string HtmlStartFragmentComment = "<!--StartFragment-->";
            const string HtmlEndFragmentComment = "<!--EndFragment-->";

            var startHtml = HtmlHeader.Length + 6 * ("0123456789".Length - "{0:D10}".Length);
            var endHtml = startHtml + Encoding.UTF8.GetByteCount(htmlString);

            var startFragment = htmlString.IndexOf(HtmlStartFragmentComment, 0);
            if (startFragment >= 0)
            {
                startFragment = startHtml + startFragment + HtmlStartFragmentComment.Length;
            }
            else
            {
                startFragment = startHtml;
            }
            var endFragment = htmlString.IndexOf(HtmlEndFragmentComment, 0);
            if (endFragment >= 0)
            {
                endFragment = startHtml + endFragment;
            }
            else
            {
                endFragment = endHtml;
            }

            // Create HTML clipboard header string
            stringBuilder.AppendFormat(HtmlHeader, startHtml, endHtml, startFragment, endFragment, startFragment, endFragment);

            // Append HTML body.
            stringBuilder.Append(htmlString);

            var returnString = stringBuilder.ToString();
            //returnString= returnString.Remove(6, 101);
            return returnString;
        }
        
        private static string RemoveClipBoardHeaderFromHtmlString(string htlmStringWithHeader)
        {
            // Example Version:1.0\nStartHTML:0000000105\nEndHTML:0000038711\nStartFragment:0000038329\nEndFragment:0000038671\n\n                             EndFragment: 0000038671
            // Version:1.0\r\nStartHTML:0000000157\r\nEndHTML:0000000185\r\nStartFragment:0000000157\r\nEndFragment:0000000185\r\nStartSelection:0000000157\r\nEndSelection:0000000185\r\n<span>REQ Description</span>"
            var startIndex = htlmStringWithHeader.IndexOf("Version:1.0");
            var endIndex1 = htlmStringWithHeader.LastIndexOf("EndSelection:") + 27;
            var endIndex2 = htlmStringWithHeader.LastIndexOf("EndFragment:") + 26;
            var endIndex = ((endIndex1 > endIndex2) ? endIndex1 : endIndex2);
            var regex = new Regex("SourceURL+.+");
            if (regex.IsMatch(htlmStringWithHeader))
            {
                // replace only first occurrence
                htlmStringWithHeader = regex.Replace(htlmStringWithHeader, string.Empty, 1);
            }
            var htmlWithoutHeader = (htlmStringWithHeader == string.Empty) ? string.Empty : htlmStringWithHeader.Remove(startIndex, endIndex);
            return htmlWithoutHeader;

        }
        #endregion

        #region Private methods
        /// <summary>
        /// Checks the start of the given token.
        /// </summary>
        /// <param name="token">Token to check.</param>
        private void ParseStartHtml(string token)
        {
            if (!token.StartsWith("StartHTML:", StringComparison.OrdinalIgnoreCase))
                throw new FormatException("Expected token: StartHTML");

            // extracts the substring after the doublepoint
            string valueString = token.Substring(token.LastIndexOf(":", StringComparison.Ordinal) + 1);
            int value;

            if (!int.TryParse(valueString, out value))
                value = -1;

            this.StartHtml = value;
        }

        /// <summary>
        /// Checks the end of the given token.
        /// </summary>
        /// <param name="token">Token to check.</param>
        private void ParseEndHtml(string token)
        {
            if (!token.StartsWith("EndHTML:", StringComparison.OrdinalIgnoreCase))
                throw new FormatException("Expected token: EndHTML");

            // extracts the substring after the doublepoint
            var valueString = token.Substring(token.LastIndexOf(":", StringComparison.Ordinal) + 1);
            int value;

            if (!int.TryParse(valueString, out value))
                value = -1;

            this.EndHtml = value;
        }
        #endregion
    }
}