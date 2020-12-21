namespace AIT.TFS.SyncService.Common.ImageComposer
{
    /// <summary>
    /// The class contains the information about one image - file and scaling of image defined in 'alt' part of 'img' token.
    /// </summary>
    public class PasteStreamImage
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PasteStreamImage"/> class.
        /// </summary>
        /// <param name="sourcePart">The source part.</param>
        /// <param name="altPart">The alt part.</param>
        public PasteStreamImage(string sourcePart, string altPart)
        {
            if (!string.IsNullOrEmpty(sourcePart))
            {
                // Set file
                File = sourcePart.Substring(5, sourcePart.Length - 6);
            }

            // Set default scales
            ScaleWidth = 100;
            ScaleHeight = 100;
            
            // Check if 'alt' part exists
            if (!string.IsNullOrEmpty(altPart))
            {
                // Extract the information
                var propsText = altPart.Substring(5, altPart.Length - 6);
                
                // Create particular properties
                var propPairs = propsText.Split(ParserDefines.ConstPropertiesDelimiter);
                foreach (var propPair in propPairs)
                {
                    // Split one property and its value
                    var parts = propPair.Split(ParserDefines.ConstPropertyPartDelimiter);
                    if (parts.Length == 2)
                    {
                        // Exact 2 parts are required
                        // 1. part is property name
                        // 2. part is property value
                        float value;
                        if (!float.TryParse(parts[1], out value))
                        {
                            // Not a value
                            continue;
                        }
                        if (parts[0] == ParserDefines.ConstScaleWidth)
                        {
                            // 'scale width' property
                            ScaleWidth = value;
                        }
                        else if (parts[0] == ParserDefines.ConstScaleHeight)
                        {
                            // 'scale height' property
                            ScaleHeight = value;
                        }
                    }
                }
            }
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the image file.
        /// </summary>
        /// <value>
        /// The image file.
        /// </value>
        public string File
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the scale of width. Default value is 100.
        /// </summary>
        /// <value>
        /// The scale of width.
        /// </value>
        public float ScaleWidth
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the scale of height. Default value is 100.
        /// </summary>
        /// <value>
        /// The scale of height.
        /// </value>
        public float ScaleHeight
        {
            get;
            private set;
        }

        #endregion Public properties
    }
}
