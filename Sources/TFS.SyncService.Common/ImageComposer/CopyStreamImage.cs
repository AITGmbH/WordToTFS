namespace AIT.TFS.SyncService.Common.ImageComposer
{
    #region Usings
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using Microsoft.Office.Interop.Word;
    #endregion

    /// <summary>
    /// The class contains the information about one image defined in one word table cell and defined in html stream on two places - vml/html part.
    /// </summary>
    public class CopyStreamImage
    {
        #region Fields
        private string _altPart;
        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CopyStreamImage"/> class.
        /// </summary>
        /// <param name="vmlFileIndex">Index of the file in vml part.</param>
        /// <param name="vmlFileLength">Length of the file in vml part.</param>
        /// <param name="htmlFileIndex">Index of the file in html part.</param>
        /// <param name="htmlFileLength">Length of the file in html part.</param>
        /// <param name="isOleObject">True if handling copy stream is a OLE object, otherwise false.</param>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "vml", Justification = "vml = vector markup language")]
        public CopyStreamImage(int vmlFileIndex, int vmlFileLength, int htmlFileIndex, int htmlFileLength, bool isOleObject)
        {
            this.VmlFileIndex = vmlFileIndex;
            this.VmlFileLength = vmlFileLength;
            this.HtmlFileIndex = htmlFileIndex;
            this.HtmlFileLength = htmlFileLength;
		    this.IsOleObject = isOleObject;
        }
        #endregion Constructors

        #region Public properties

        /// <summary>
        /// True if the handling <see cref="CopyStreamImage"/> is an OLE object, otherwise false.
        /// </summary>
        public bool IsOleObject { get; private set; }

        /// <summary>
        /// Gets the index of the file in vml part.
        /// </summary>
        /// <value>
        /// The index of the file in vml part.
        /// </value>
        public int VmlFileIndex
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the length of the file in vml part.
        /// </summary>
        /// <value>
        /// The length of the file in vml part.
        /// </value>
        public int VmlFileLength
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the index of the file in html part.
        /// </summary>
        /// <value>
        /// The index of the file in html part.
        /// </value>
        public int HtmlFileIndex
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the length of the file in html part.
        /// </summary>
        /// <value>
        /// The length of the file in html part.
        /// </value>
        public int HtmlFileLength
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the start index of the position of whole 'alt="..."' part. If 'alt' part exists.
        /// If not exists, index means the position, where should be inserted new 'alt' part.
        /// Note: Insert a blank after new inserted part.
        /// </summary>
        /// <value>
        /// The position where old 'alt' part starts, or the position where should be new 'alt' part inserted.
        /// </value>
        public int HtmlAltStart
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the end index of the position of whole 'alt="..."' part - index points to the second quotation marks.
        /// Value of -1 means that the 'alt="..."' don't exists at all.
        /// </summary>
        /// <value>
        /// The position where the old 'alt' part ends, or the information that 'alt' part don't exists at all (-1 value).
        /// </value>
        public int HtmlAltEnd
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the alt part for the image where is stored the information about image scaling.
        /// </summary>
        /// <value>
        /// The alt part with scaling information.
        /// </value>
        public string AltPart
        {
            get
            {
                // MIS: exception for OLE Objects, because they are handled as images and dont have inline shapes therefore altpart is not set. 
                // Therefore set it manually to max hieght and size
                if (string.IsNullOrEmpty(_altPart))
                {
                    SetImageInformation(100,100);
                }

                return _altPart;
            }
            private set
            {
                _altPart = value;
            }
        }

        #endregion Public properties

        #region Public methods

        /// <summary>
        /// The method sets the <see cref="AltPart"/> property based on <see cref="InlineShape"/>.
        /// </summary>
        /// <param name="inlineShape">The inline shape to use for creating of <see cref="AltPart"/>.</param>
        internal void SetImageInformation(InlineShape inlineShape)
        {
            SetImageInformation(inlineShape.ScaleWidth, inlineShape.ScaleHeight);
        }

        /// <summary>
        /// Sets the image <see cref="AltPart"/> property based on width and size.
        /// </summary>
        /// <param name="scaleWidth">The scale width of the image.</param>
        /// <param name="scaleHeight">The scale hieght of the image.</param>
        private void SetImageInformation(float scaleWidth, float scaleHeight)
        {
            this.AltPart = string.Format(
                CultureInfo.InvariantCulture,
                "alt=\"{0}{1}{2}{3}{4}{5}{6}\"{7}",
                ParserDefines.ConstScaleWidth,
                ParserDefines.ConstPropertyPartDelimiter,
                scaleWidth,
                ParserDefines.ConstPropertiesDelimiter,
                ParserDefines.ConstScaleHeight,
                ParserDefines.ConstPropertyPartDelimiter,
                scaleHeight,
                HtmlAltEnd == -1 ? " " : string.Empty);
        }

        #endregion Public methods
    }
}
