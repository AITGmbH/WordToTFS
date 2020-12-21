namespace AIT.TFS.SyncService.Common.ImageComposer
{
    #region Usings
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    using Microsoft.Office.Interop.Word;
    using AIT.TFS.SyncService.Factory;
    using System.Globalization;
    using AIT.TFS.SyncService.Common.Properties;
    using System.Runtime.InteropServices;
    #endregion

    /// <summary>
    /// Parser to find every image file definied in vml and html part of html stream created by copy of word table cell content.
    /// </summary>
    public class PasteStreamParser
    {
        #region Private fields

        private List<PasteStreamImage> imageFiles = new List<PasteStreamImage>();

        #endregion Private fields

        #region Private properties

        /// <summary>
        /// Gets all found image files.
        /// </summary>
        /// <value>
        /// The found image files.
        /// </value>
        public IList<PasteStreamImage> ImageFiles
        {
            get
            {
                return this.imageFiles;
            }
        }

        #endregion Private properties

        #region Public methods

        /// <summary>
        /// Parses the specified stream and find images.
        /// </summary>
        /// <param name="stream">The stream to parse and find images.</param>
        /// <param name="range">Corresponding range in word.</param>
        public void ParseAndRepairAfterPaste(string stream, Range range)
        {
            if (range == null)
            {
                throw new ArgumentNullException("range");
            }

            Parse(stream);

            // Don't access the shapes with '[]' operator
            var shapeIndex = 0;
            foreach (InlineShape shape in range.InlineShapes)
            {
                try
                {
                    // Check if the element exists and use it
                    if (shapeIndex < this.imageFiles.Count)
                    {
                        shape.ScaleWidth = this.imageFiles[shapeIndex].ScaleWidth;
                        shape.ScaleHeight = this.imageFiles[shapeIndex].ScaleHeight;
                    }
                    else
                    {

                        shape.ScaleWidth = 100;
                        shape.ScaleHeight = 100;
                    }
                }
                catch (COMException ex)
                {
                    if (ex.ErrorCode == -2146823595)
                    {
                        // <hr /> appear as shapes but ScaleWidth and ScaleHeight cannot be set.
                        // The following exception is thrown: "This member cannot be accessed on a horizontal line."
                        SyncServiceTrace.W(string.Format(CultureInfo.CurrentCulture, Resources.Shape_Scale100_Failed, ex.Message));
                    }
                    else
                    {
                        throw;
                    }
                }
                
                shapeIndex++;
            }
        }

        #endregion Public methods

        #region Private methods

        /// <summary>
        /// The method parses the specified stream and finds images.
        /// </summary>
        /// <param name="stream">The stream to parse and find images.</param>
        private void Parse(string stream)
        {
            // Clear list
            this.imageFiles.Clear();

            // Find 'img' parts with image
            var pattern = @"<img[^>]*>";
            var imgMatches = Regex.Matches(stream, pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match img in imgMatches)
            {
                // Find src part
                var srcMatches = Regex.Matches(img.Value, @"src=\""[^\""]*\""", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                if (srcMatches.Count == 1)
                {
                    var src = srcMatches[0].Value;
                    var alt = string.Empty;

                    // Find alt part
                    var altMatches = Regex.Matches(img.Value, @"alt=\""[^\""]*\""", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                    if (altMatches.Count == 1)
                    {
                        alt = altMatches[0].Value;
                    }
                    
                    // Create class for this image
                    this.imageFiles.Add(new PasteStreamImage(src, alt));
                }
            }
        }

        #endregion Private methods
    }
}
