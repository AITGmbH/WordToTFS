namespace AIT.TFS.SyncService.Common.Helper
{
    #region Usings
    using System;
    using System.IO;
    using System.Text.RegularExpressions;
    #endregion

    /// <summary>
    /// Class implements a wrapper of one html image tag. Class used to replace the file property in html image tag.
    /// </summary>
    public sealed class HtmlImage : IDisposable
    {
        #region Fields
        private readonly string _tag;
        private readonly DirectoryInfo _tempdir;
        private readonly bool _deleteTempFiles;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="HtmlImage"/> class.
        /// </summary>
        /// <param name="imageTag">Tag of the image to wrap in this class.</param>
        /// <param name="tempdir">Directory to use for the temporary files.</param>
        /// <param name="deleteTempFiles">Sets whether to delete temporary files on dispose.</param>
        internal HtmlImage(string imageTag, DirectoryInfo tempdir, bool deleteTempFiles)
        {
            this._tag = imageTag;
            this._tempdir = tempdir;
            this._deleteTempFiles = deleteTempFiles;

            // extract the file path --> its between the src=" and the next closing "
            var start = imageTag.IndexOf("src=\"", StringComparison.OrdinalIgnoreCase) + 5;
            var end = imageTag.IndexOf("\"", start + 1, StringComparison.OrdinalIgnoreCase);

            this.Source = imageTag.Substring(start, end - start);
            this.Uri = new Uri(this.Source);
            this.LocalFileInfo = new FileInfo(this.Uri.AbsolutePath);
        }
        #endregion

        #region Properties

        /// <summary>
        /// The value of the source attribute
        /// </summary>
        public string Source
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the original image tag.
        /// </summary>
        public string OriginalImageTag
        {
            get
            {
                return this._tag;
            }
        }

        /// <summary>
        /// Gets whether the attachment filename if this image is a link to the attachment handler
        /// </summary>
        public string FileNameQueryParameter
        {
            get
            {
                var filename = Regex.Match(this._tag, "FileName=(?<Filename>[^&\"]+)");
                var captures = filename.Groups["Filename"].Captures;
                return captures.Count > 0 ? captures[0].Value : string.Empty;
            }
        }

        /// <summary>
        /// Gets the new (created) image tag.
        /// </summary>
        public string UpdatedImageTag
        {
            get
            {
                // Replace // with /
                var uri = Regex.Replace(this.Uri.AbsoluteUri, @"(?<!:)\/\/", "/");

                // FIX for <v:imagedata into <img tag replacement
                var fullTag = this._tag.Replace("<v:imagedata", "<img");
                fullTag = fullTag.Replace("o:Title=", "alt=");

                return fullTag.Replace(this.Source, uri);
            }
        }

        /// <summary>
        /// Gets the file info of the local file.
        /// </summary>
        public FileInfo LocalFileInfo
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the local file name;
        /// </summary>
        public string Name
        {
            get
            {
                return LocalFileInfo.Name;
            }
            set
            {
                if (null != value && LocalFileInfo.Name != value)
                {
                    if (!_tempdir.Exists)
                    {
                        _tempdir.Create();
                    }

                    var destination = Path.Combine(_tempdir.FullName, value);

                    if (File.Exists(destination))
                    {
                        File.Delete(destination);
                    }

                    LocalFileInfo = LocalFileInfo.CopyTo(destination);
                    Uri = new Uri(LocalFileInfo.FullName);
                }
            }
        }

        /// <summary>
        /// Uri of this image
        /// </summary>
        public Uri Uri
        {
            get;
            set;
        }

        #endregion

        #region Public methods
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_deleteTempFiles && null != LocalFileInfo && LocalFileInfo.Exists)
            {
                LocalFileInfo.Delete();
                LocalFileInfo = null;
            }
        }
        #endregion
    }
}