namespace AIT.TFS.SyncService.Common.Helper
{
    #region Usings
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using HtmlAgilityPack;
    #endregion

    /// <summary>
    /// Helper class implements helper functionality for the work with images in html stream.
    /// </summary>
    public sealed class HtmlHelper : IDisposable
    {
        #region Fields

        private static readonly Regex _imageTagExpression = new Regex(@"(<(img\s|v:imagedata\ssrc)[^>]*>)", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);
        // private static Regex imageFindExpression = new Regex("<(img|v:imagedata)[^>]src=\"file://[^\"].(png|gif|jpg|jpeg|bmp)", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);

        // MIK 17.2.2016 - we are using 'alt' part to store the scaling of image. Don't remove this part
        // Old pattern: (<u\d+:p>(&nbsp;)?</u\d+:p>)|(alt=""([^\""]*)\"")|(v:shapes=""([^\""]*)\"")
        // New pattern: (<u\d+:p>(&nbsp;)?</u\d+:p>)|(v:shapes=""([^\""]*)\"")
        private static readonly Regex _tagExpression = new Regex(@"(<u\d+:p>(&nbsp;)?</u\d+:p>)|(v:shapes=""([^\""]*)\"")", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);

        private readonly string _stream;
        private IList<HtmlImage> _images;
        private DirectoryInfo _tempDirectory;
        private readonly bool _deleteTempFiles;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="HtmlHelper"/> class.
        /// </summary>
        /// <param name="stream">HTML stream to process.</param>
        /// <param name="deleteTempFiles">Sets whether to delete temp files.</param>
        public HtmlHelper(string stream, bool deleteTempFiles)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");

            _deleteTempFiles = deleteTempFiles;
            _stream = Clean(stream);
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Gets the changed html stream.
        /// </summary>
        public string Html
        {
            get
            {
                string stream = _stream;

                foreach (HtmlImage image in Images)
                    stream = stream.Replace(image.OriginalImageTag, image.UpdatedImageTag);

                stream = UpdateUnsortedLists(stream);

                return stream;
            }
        }

        /// <summary>
        /// Gets the list of all images (<see cref="HtmlImage"/>) in the html stream.
        /// </summary>
        public IList<HtmlImage> Images
        {
            get
            {
                if (null == _images)
                {
                    _images = new List<HtmlImage>();

                    foreach (Match match in _imageTagExpression.Matches(_stream))
                    {
                        var image = new HtmlImage(match.Value, TempDirectory, _deleteTempFiles);
                        _images.Add(image);
                    }
                }
                return _images;
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Removes the paragraph and division tags.
        /// </summary>
        /// <param name="htmlContent">The HTML content string.</param>
        /// <returns></returns>
        public static string ReplaceParagraphAndDivisionTags(string htmlContent)
        {
            /* htmlContent from MTM is different as htmlContent from web - MTM uses paragraphes for each line 
             * 
             * MTM (using Alt+Enter for linebreaks, Not available in Web): 
             * <DIV><P>Generated in MTM with line breaks</P><P>Line1</P><P>Line2</P><P>Line3</P></DIV> 
             * 
             * Web (using Shift+Enter for linebreaks, not available in MTM): 
             * <DIV><P>Using Web<BR />Line1<BR />Line2<BR />Line3 </P></DIV> 
             * 
             * Ergo: Web -> Single <P>-Tag, MTM -> <p>-Tag foreach linebreak --> linebreaks from MTM are not shown in Word 
             * Solution: Before all other replacements are done, replace all </p>-Tags by <br/> for line breaks */
            var text = Regex.Replace(htmlContent, @"(</p>)", @"<br/>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(<p( [^>]*)*>)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(<p>)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(</p( [^>]*)*>)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(</p>)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);

            text = Regex.Replace(text, @"(<div( [^>]*)*>)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(<div>)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(</div( [^>]*)*>)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(</div>)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);

            return text;
        }

        /// <summary>
        /// Extracts the plain text from HTML.
        /// </summary>
        /// <param name="htmlContent">Content of the HTML.</param>
        /// <returns>Plain text extracted from HTML.</returns>
        public static string ExtractPlaintext(string htmlContent)
        {
            if (htmlContent == null) throw new ArgumentNullException("htmlContent");

            var body = htmlContent.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            var endBody = htmlContent.IndexOf("</body>", StringComparison.OrdinalIgnoreCase);

            if (body > 0 && endBody > body)
            {
                htmlContent = htmlContent.Substring(body, endBody - body);
            }
            else
            {
                // when editing html code from word in tfs, head tags are removed but the style tags remain and would appear as "extracted plaintext"
                htmlContent = Regex.Replace(htmlContent, @"<style>.*?</style>", string.Empty, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            }

            // Replace all opening and closing paragraph tags (with and without attributes) by spaces.
            var text = Regex.Replace(htmlContent, @"(<p( [^>]*)*>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(<p>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(</p( [^>]*)*>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(</p>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // <br> tags in TFS are translated to <o:p>&nbsp;</o:p> in Word, so replace <br> tags by non breaking spaces.
            text = Regex.Replace(text, @"(<br>)", @"&nbsp;", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Replace all list item tags (with and without attributes) by spaces.
            text = Regex.Replace(text, @"(<li( [^>]*)*>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(<li>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Replace all opening and closing heading tags (with and without attributes) by spaces.
            text = Regex.Replace(text, @"(<h(\d+)( [^>]*)*>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(<h(\d+)>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(</h(\d+)( [^>]*)*>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(</h(\d+)>)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // This regex is "[\s] without the actual space character"
            var regexSpace = new Regex(@"[\f\n\r\t\v]+");
            text = regexSpace.Replace(text, " ");

            // Remove all comment or conditional tags and their content.
            text = Regex.Replace(text, @"(<!--(.+?)-->)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Remove all remaining tags.
            text = Regex.Replace(text, @"(<[^>]+>)", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Add the extra spaces defined by non breaking spaces.
            text = Regex.Replace(text, @"(&nbsp;)|(&#160;)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            //// Replace all consecutive spaces by just a single space.
            text = Regex.Replace(text, @"(\s+)", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);


            text = text.Trim();

            return text;
        }

        /// <summary>
        /// The given html string is searched for image tags and shrinks the width and height of the images to the given percentage.
        /// </summary>
        /// <param name="htmlContent">Content of the HTML string.</param>
        /// <param name="percentage">Percentage is a Key-Value-Store. The keys are the image IDs. The width and height of an image is multiplied by the corresponding value.</param>
        /// <returns>The html string with the width and height replaced by the shrunken values.</returns>
        public static string ShrinkImages(string htmlContent, Dictionary<string, float> percentage)
        {
            const string IMAGE_TAG_PATTERN = @"<img ([^>]+?)>";
            var matches = Regex.Matches(htmlContent, IMAGE_TAG_PATTERN, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                foreach (string key in percentage.Keys)
                {
                    if (match.Value.Contains(key))
                    {
                        var replacement = match.Value;
                        replacement = ReplaceImageSize(percentage[key], replacement, "width");
                        replacement = ReplaceImageSize(percentage[key], replacement, "height");
                        htmlContent = htmlContent.Replace(match.Value, replacement);
                    }
                }
            }
            return htmlContent;
        }

        /// <summary>
        /// Word does not show images if the HTML does not contain both a width and a height attribute. Therefore if only one of both exists it is necessary to add the missing one. The value for the missing attribute is calculated based on the ratio of the width and height element within the style-attribute.
        /// </summary>
        /// <param name="htmlContent">Html-String</param>
        /// <returns>Html-String in which missing image size attributes have been added.</returns>
        public static string AddMissingImageSizeAttributes(string htmlContent)
        {
            const string IMAGE_TAG_PATTERN = @"<img ([^>]+?)>";
            var matches = Regex.Matches(htmlContent, IMAGE_TAG_PATTERN, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                var widthExists = Regex.IsMatch(match.Value, @"width=");
                var heightExists = Regex.IsMatch(match.Value, @"height=");

                var replacement = string.Empty;
                if (!widthExists)
                {
                    replacement = GetHtmlWithNewSizeAttribute(match.Value, @"width");
                }
                else if (!heightExists)
                {
                    replacement = GetHtmlWithNewSizeAttribute(match.Value, @"height");
                }

                // if anything has been replaced, write the new value into htmlContent
                if (replacement != string.Empty && replacement != match.Value)
                {
                    htmlContent = htmlContent.Replace(match.Value, replacement);
                }
            }
            return htmlContent;
        }

        /// <summary>
        /// The given html string is searched for image tags and shrinks the width and height of the images to the given percentage.
        /// </summary>
        /// <param name="htmlContent">Content of the HTML string.</param>
        /// <param name="percentage">The width and height of all images within that html value are multiplied by the percentage value.</param>
        /// <returns>The html string with the width and height replaced by the shrunken values.</returns>
        public static string ShrinkImagesAllSamePercentage(string htmlContent, float percentage)
        {
            const string IMAGE_TAG_PATTERN = @"<img ([^>]+?)>";
            var matches = Regex.Matches(htmlContent, IMAGE_TAG_PATTERN, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                var replacement = match.Value;
                replacement = ReplaceImageSize(percentage, replacement, "width");
                replacement = ReplaceImageSize(percentage, replacement, "height");
                htmlContent = htmlContent.Replace(match.Value, replacement);
            }
            return htmlContent;
        }

        #endregion

        #region Private methods

        private static string GetHtmlWithNewSizeAttribute(string imageHtml, string attributeName)
        {
            // if nothing will be replaced, the original imageHtml text will be returned
            var replacement = imageHtml;

            if (Regex.IsMatch(imageHtml, @"width:([^>]+?)px") && Regex.IsMatch(imageHtml, @"height:([^>]+?)px"))
            {
                var styleWidthAttribute = (Regex.Matches(imageHtml, @"width:([^>]+?)px"))[0].Value;
                var styleWidthAttributeDouble = Convert.ToDouble(styleWidthAttribute.Replace("width:", "").Replace("px", ""));

                var styleHeightAttribute = (Regex.Matches(imageHtml, @"height:([^>]+?)px"))[0].Value;
                var styleHeightAttributeDouble = Convert.ToDouble(styleHeightAttribute.Replace("height:", "").Replace("px", ""));

                var ratio = styleWidthAttributeDouble / styleHeightAttributeDouble;

                var sizeAttributeName = (attributeName == "width") ? "height" : "width";
                var regex = string.Format(@"{0}{1}", sizeAttributeName, "=\\D*(?<digit>\\d+)");// height=\\D*(?<digit>\\d+)

                if (Regex.IsMatch(imageHtml, regex))
                {
                    var sizeAttributeDouble = Convert.ToDouble(Regex.Matches(imageHtml, regex)[0].Groups["digit"].Value);

                    var newSizeAttributeInt = 0;
                    var newSizeAttribute = string.Empty;
                    if (attributeName == "width")
                    {
                        newSizeAttributeInt = (int)Math.Round(sizeAttributeDouble * ratio, 0);
                        newSizeAttribute = "width=" + newSizeAttributeInt;
                    }
                    else
                    {
                        newSizeAttributeInt = (int)Math.Round(sizeAttributeDouble / ratio, 0);
                        newSizeAttribute = "height=" + newSizeAttributeInt;
                    }

                    replacement = imageHtml.Replace("<img", @"<img " + newSizeAttribute);
                }
            }

            return replacement;
        }

        private static string ReplaceImageSize(float percentage, string replacement, string searchString)
        {
            var matches = Regex.Matches(replacement, searchString + @"(\s*?)[=:](.*?)(\d+)", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                int value;
                if (int.TryParse(match.Groups[3].Value, out value))
                {
                    value = (int)Math.Round(percentage * value);
                    var widthReplacement = Regex.Replace(match.Value, match.Groups[3].Value, value.ToString(CultureInfo.InvariantCulture));
                    replacement = Regex.Replace(replacement, match.Value, widthReplacement);
                }
            }
            return replacement;
        }

        private static string Clean(string rawHtml)
        {
            var newHtml = new StringBuilder(rawHtml);
            // remove <!--[if gte svml....<![endif]--> tag content which contains <v:shape tag.
            // this fix is necessary for web access regarding the bitmaps and shapes
            var imageDataPattern = @"<!--\[if\sgte\svml(.+?)<!\[endif]-->";
            var matches = Regex.Matches(rawHtml, imageDataPattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                if (match.Value.IndexOf("<v:shape", StringComparison.OrdinalIgnoreCase) > 0)
                {
                    newHtml.Replace(match.Value, string.Empty);
                }
            }

            // remove <!--[if gte svml....<![endif]--> tag content which contains <v:shape tag.
            // this fix is necessary for web access regarding the equation
            var equationPattern = @"<!--\[if\sgte\smsEquation(.+?)<!\[endif]-->";
            matches = Regex.Matches(newHtml.ToString(), equationPattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                newHtml.Replace(match.Value, string.Empty);
            }

            // remove border tags for images
            // this fix is necessary for web access
            var imagePattern = @"<!\[if\s!vml]>(.+?)<!\[endif]>";
            var imageReplacePattern = @"(<!\[if\s!vml]>)|(<!\[endif]>)";
            matches = Regex.Matches(newHtml.ToString(), imagePattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                if (match.Value.IndexOf("<img", StringComparison.OrdinalIgnoreCase) > 0)
                {
                    var newTag = Regex.Replace(match.Value, imageReplacePattern, string.Empty,
                                               RegexOptions.Singleline | RegexOptions.IgnoreCase);
                    newHtml.Replace(match.Value, newTag);
                }
            }

            // remove border tags for msEquation
            // this fix is necessary for web access
            var notEquationPattern = @"<!\[if\s!msEquation]>(.+?)<!\[endif]>";
            var equationReplacePattern = @"(<!\[if\s!msEquation]>)|(<!\[endif]>)";
            matches = Regex.Matches(newHtml.ToString(), notEquationPattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                var newTag = Regex.Replace(match.Value, equationReplacePattern, string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);
                newHtml.Replace(match.Value, newTag);
            }

            // remove such <u6:p></u6:p> patterns to avoid revision number change for field with html content
            // When the TFS HTML field is saved into Word document field then such <u6:p></u6:p> tag is added
            // into html content with increased number and this cause that the html content in word is different
            // from TFS field html content even if nothing was changed.
            matches = _tagExpression.Matches(newHtml.ToString());
            foreach (Match match in matches)
            {
                newHtml.Replace(match.Value, string.Empty);
            }

            return newHtml.ToString();
        }

        /// <summary>
        /// Modifies the html string to bypass the display bug in visual studio when displaying html lists in work items.
        /// </summary>
        /// <param name="html">Unmodified html string.</param>
        /// <returns>Fixed html string.</returns>
        private static string UpdateUnsortedLists(string html)
        {
            html = html.Replace(@"<p class=MsoListParagraphCxSpFirst style='",
                                @"<p class=MsoListParagraphCxSpFirst style='margin-left:36.0pt;");
            html = html.Replace(@"<p class=MsoListParagraphCxSpMiddle style='",
                                @"<p class=MsoListParagraphCxSpMiddle style='margin-left:36.0pt;");
            html = html.Replace(@"<p class=MsoListParagraphCxSpLast style='",
                                @"<p class=MsoListParagraphCxSpLast style='margin-left:36.0pt;");

            var doc = new HtmlDocument();
            doc.OptionAutoCloseOnEnd = true;
            doc.LoadHtml(html);

            RemoveStyleinformationFromBulletPoints(doc.DocumentNode);
            html = doc.DocumentNode.WriteTo();

            return html;
        }


        /// <summary>
        /// ReSharper disable once CSharpWarnings::CS1570
        /// Remove the style information from span tags with font family "Time new Roman" followed by &nbsp; this is caused by the conversion of UL to something that looks like a list
        /// </summary>
        /// <param name="node"></param>
        private static void RemoveStyleinformationFromBulletPoints(HtmlNode node)
        {
            var xpaths = IdentifyNodesWithBulletPoints(node);

            foreach (string xpath in xpaths)
            {
                var bulletPointNode = node.SelectSingleNode(xpath);

                // Remove the attributes
                bulletPointNode.Attributes["style"].Remove();

                // Replace the unnecessary blanks with a single one
                if (string.IsNullOrWhiteSpace(HtmlEntity.DeEntitize(bulletPointNode.InnerHtml)))
                {
                    bulletPointNode.InnerHtml = "&nbsp;";
                }
            }
        }

        /// <summary>
        /// This methods searches all nodes that have a span tag with times new roman and multiple blanks
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        private static List<string> IdentifyNodesWithBulletPoints(HtmlNode node)
        {
            List<string> xpaths = new List<string>();

            if (node.Name.Contains("span"))
            {
                foreach (var nodethatmatches in node.Attributes.Where(x => x.Name.Contains("style") && x.Value.Contains("Times New Roman")))
                {

                    if (string.IsNullOrWhiteSpace(HtmlEntity.DeEntitize(node.InnerHtml)))
                    {
                        xpaths.Add(nodethatmatches.XPath);
                    }
                }
            }

            foreach (var childNode in node.ChildNodes)
            {
                xpaths.AddRange(IdentifyNodesWithBulletPoints(childNode));
            }

            return xpaths;
        }

        #endregion

        #region Private properties

        /// <summary>
        /// Gets the temporary directory. Temporary directory isn't created.
        /// </summary>
        private DirectoryInfo TempDirectory
        {
            get
            {
                if (null == _tempDirectory)
                {
                    string directory = Path.Combine(TempFolder.CreateNewTempFolder("HtmlHelper"), Guid.NewGuid().ToString());
                    _tempDirectory = new DirectoryInfo(directory);
                }

                return _tempDirectory;
            }
        }

        #endregion

        #region IDisposable members

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (null != _images)
            {
                foreach (HtmlImage image in Images)
                {
                    image.Dispose();
                }
            }

            if (_deleteTempFiles && TempDirectory.Exists)
            {
                TempDirectory.Delete(true);
            }
        }

        #endregion
    }
}
