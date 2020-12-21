#region Usings
using System;
using System.Linq;
using AIT.TFS.SyncService.Common.Helper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
#endregion

namespace TFS.SyncService.Common.Test.Unit
{

    /// <summary>
    /// Unit tests for the HTML Helper
    /// </summary>
    [TestClass]
    public class HTMLHelperTestUnit
    {
        #region Consts
        private const string FirstVerifcationString = "This string is used for verification purposes. It is the first of two strings";
        private const string SecondVerifcationString = "This string is used for verification purposes. It is the second of two strings";

        private const string ParagraphWithTimesNewRomanAndWhiteSpaces = "<span lang=EN style='font-family:\"Times New Roman\",serif'>This is a paragraph formatted in Times New Roman. It contains the verification string: " + FirstVerifcationString + ". In addition it has multiple whitespaces.StartofWhiteSpaces         End of whitespaces.</span>";
        private const string ParagraphWithTimesNewRomanAndNBSP = "<span lang=EN style='font-family:\"Times New Roman\",serif'>This is a paragraph formatted in Times New Roman. It contains the verification string: " + FirstVerifcationString + ". In addition it has multiple NBSP.StartofNBSPs&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;\r\n&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbspEnd of NBSP.</span>";

        private const string BulletPointListFollowedByText = "<p class=RequireTableDescription><span lang=EN style='font-family:\"Times New Roman\",serif'>This is a paragraph with a bullet point list, not followed by text" + FirstVerifcationString + " </span></p><p class=RequireTableDescription style='margin-left:36.0pt;text-indent:-18.0pt'><span lang=EN style='font-family:Symbol'>·<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><span lang=EN style='font-family:\"Times New Roman\",serif'>B1 " + SecondVerifcationString + "</span></p><p class=RequireTableDescription style='margin-left:36.0pt;text-indent:-18.0pt'><span lang=EN style='font-family:Symbol'>·<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span><span lang=EN style='font-family:\"Times New Roman\",serif'>B2</span></p><p class=RequireTableDescription style='margin-left:36.0pt;text-indent:-18.0pt'><span lang=EN style='font-family:Symbol'>·<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span><span lang=EN style='font-family:\"Times New Roman\",serif'>&nbsp;</span></p><p class=RequireTableDescription><span lang=EN style='font-family:\"Times New Roman\",serif'>End of List</span></p>";
        private const string BulletPointListNotFollowedByText = "<p class=RequireTableDescription><span lang=EN style='font-family:\"Times New Roman\",serif'>This is a paragraph with a bullet point list, not followed by text" + FirstVerifcationString + " </span></p><p class=RequireTableDescription style='margin-left:36.0pt;text-indent:-18.0pt'><span lang=EN style='font-family:Symbol'>·<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><span lang=EN style='font-family:\"Times New Roman\",serif'>B1 " + SecondVerifcationString + "</span></p><p class=RequireTableDescription style='margin-left:36.0pt;text-indent:-18.0pt'><span lang=EN style='font-family:Symbol'>·<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span><span lang=EN style='font-family:\"Times New Roman\",serif'>B2</span></p><p class=RequireTableDescription style='margin-left:36.0pt;text-indent:-18.0pt'><span lang=EN style='font-family:Symbol'>·<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span><span lang=EN style='font-family:\"Times New Roman\",serif'>&nbsp;</span></p><p class=RequireTableDescription><span lang=EN style='font-family:\"Times New Roman\",serif'>&nbsp;</span></p>";
        #endregion

        #region TestMethods
        /// <summary>
        /// Tests that the whitespaces in a closed bullet point list are removed.
        /// </summary>
        [TestMethod]
        public void HTMLHelperUnitTest_TestHTMLRetrieval_TestRemovalOfSpanTagsAndWhitespacesFromBulletPointListNotFollowedByText_ShouldRemoveSpanTagFromBulletPointList()
        {
            var htmlHelper = new HtmlHelper(BulletPointListNotFollowedByText, true);
            var valueAfterHtmlParsing = htmlHelper.Html;

            Assert.IsTrue(valueAfterHtmlParsing.Contains(FirstVerifcationString));
            Assert.IsTrue(valueAfterHtmlParsing.Contains(SecondVerifcationString));
            //Should not contains more than one nbsp;
            Assert.IsFalse(valueAfterHtmlParsing.Contains("&nbsp;&nbsp;"));
            //Should contain single nbsp;
            Assert.IsTrue(valueAfterHtmlParsing.Contains("&nbsp"));
        }

        /// <summary>
        /// Tests that the whitespaces in a unclosed bullet point list are removed.
        /// </summary>
        [TestMethod]
        public void HTMLHelperUnitTest_TestHTMLRetrieval_TestRemovalOfSpanTagsAndWhitespacesFromBulletPointListFollowedByText_ShouldRemoveSpanTagFromBulletPointList()
        {
            var htmlHelper = new HtmlHelper(BulletPointListFollowedByText, true);
            var valueAfterHtmlParsing = htmlHelper.Html;

            Assert.IsTrue(valueAfterHtmlParsing.Contains(FirstVerifcationString));
            Assert.IsTrue(valueAfterHtmlParsing.Contains(SecondVerifcationString));
            //Should not contains more than one nbsp;
            Assert.IsFalse(valueAfterHtmlParsing.Contains("&nbsp;&nbsp;"));
            //Should contain single nbsp;
            Assert.IsTrue(valueAfterHtmlParsing.Contains("&nbsp"));
        }

        /// <summary>
        /// Tests that a paragraph formatted in times new roman with multiple whitespaces is not removed
        /// </summary>
        [TestMethod]
        public void HTMLHelperUnitTest_TestHTMLRetrieval_TestRemovalOfSpanTagsAndWhitespacesFromBulletPointLists_ShouldNotRemoveParagraphsWithTimesNewRoman()
        {
            var htmlHelper = new HtmlHelper(ParagraphWithTimesNewRomanAndWhiteSpaces, true);

            var valueAfterHtmlParsing = htmlHelper.Html;
            Assert.IsTrue(ParagraphWithTimesNewRomanAndWhiteSpaces.Contains(FirstVerifcationString));

            //Should contains more than one nbsp;
            Assert.IsFalse(valueAfterHtmlParsing.Contains("&nbsp;&nbsp;"));
            //Should contain single nbsp;
            Assert.IsFalse(valueAfterHtmlParsing.Contains("&nbsp"));

            //Should contain single whitespaces;
            Assert.IsTrue( valueAfterHtmlParsing.Any ( Char.IsWhiteSpace));
        }

        /// <summary>
        /// Tests that a paragraph formatted in times new roman with multiple NBSP is not removed
        /// </summary>
        [TestMethod]
        public void HTMLHelperUnitTest_TestHTMLRetrieval_TestRemovalOfSpanTagsAndNBSPFromBulletPointLists_ShouldNotRemoveParagraphsWithTimesNewRoman()
        {
            var htmlHelper = new HtmlHelper(ParagraphWithTimesNewRomanAndNBSP, true);

            var valueAfterHtmlParsing = htmlHelper.Html;
            Assert.IsTrue(ParagraphWithTimesNewRomanAndNBSP.Contains(FirstVerifcationString));
            //Should contains more than one nbsp;
            Assert.IsTrue(valueAfterHtmlParsing.Contains("&nbsp;&nbsp;"));
            //Should contain single nbsp;
            Assert.IsTrue(valueAfterHtmlParsing.Contains("&nbsp"));
        }
        #endregion

    }
}
