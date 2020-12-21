using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Net;

namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects
{
    using AIT.TFS.SyncService.Common.Helper;

    /// <summary>
    /// Class implements the attachment wrapper functionality.
    /// </summary>
    internal class TfsAttachment : IDisposable
    {
        private readonly Attachment _attachment;
        private readonly TfsWorkItem _workItem;
        private FileInfo _localFile;

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsAttachment"/> class.
        /// </summary>
        /// <param name="workItem">Work item that holds the attachment.</param>
        /// <param name="attachment">Attachment to process.</param>
        internal TfsAttachment(TfsWorkItem workItem, Attachment attachment)
        {
            if (workItem == null)
                throw new ArgumentNullException("workItem");

            if (attachment == null)
                throw new ArgumentNullException("attachment");

            _workItem = workItem;
            _attachment = attachment;
        }

        #region IDisposable Members

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (null != _localFile && _localFile.Exists)
            {
                _localFile.Delete();
                _localFile = null;
            }
        }

        #endregion

        /// <summary>
        /// Method compares the associated attachment with file.
        /// </summary>
        /// <param name="info">File to compare with the associated attachment.</param>
        /// <returns>True if the given file is identical with the associated attachment, otherwise false.</returns>
        public bool Compare(FileInfo info)
        {
            Download();
            return FileCompare(_localFile.FullName, info.FullName);
        }

        /// <summary>
        /// Download file from Web uri.
        /// </summary>
        /// <param name="uri">Uri where the file is located.</param>
        /// <param name="credentials">Download credentials.</param>
        /// <param name="fileName">File name with path where to file save.</param>
        [SuppressMessage("Microsoft.Reliability", "CA2000:DisposeObjectsBeforeLosingScope", Justification = "Message targets public property of web client which is disposed")]
        public static void DownloadWebFile(Uri uri, ICredentials credentials, string fileName)
        {
            using (var request = new WebClient { Credentials = credentials })
            {
                request.DownloadFile(uri, fileName);
            }
        }

        /// <summary>
        /// Download the associated work item attachment.
        /// </summary>
        private void Download()
        {
            if (null == _localFile)
            {
                _localFile = new FileInfo(TempFolder.CreateTemporaryFile());
                DownloadWebFile(_attachment.Uri, _workItem.WorkItem.Store.TeamProjectCollection.Credentials,
                                _localFile.FullName);
            }
        }

        /// <summary>
        /// Method compares content of two files. Files are considered equal if they have the same name OR the same content
        /// </summary>
        /// <param name="file1">File one to compare.</param>
        /// <param name="file2">File two to compare.</param>
        /// <returns>True if the files are have the same name or are identical, otherwise false.</returns>
        internal static bool FileCompare(string file1, string file2)
        {
            var file1Byte = 0;
            var file2Byte = 0;

            // Determine if the same file was referenced two times.
            if (file1 == file2)
            {
                return true;
            }

            // Open the two files.
            using (FileStream fs1 = new FileStream(file1, FileMode.Open), fs2 = new FileStream(file2, FileMode.Open))
            {

                // Check the file sizes. If they are not the same, the files 
                // are not the same.
                if (fs1.Length != fs2.Length)
                {
                    return false;
                }

                // Read and compare a byte from each file until either a
                // non-matching set of bytes is found or until the end of
                // file1 is reached (files have the same length at this point).
                do
                {
                    file1Byte = fs1.ReadByte();
                    file2Byte = fs2.ReadByte();
                } while ((file1Byte == file2Byte) && (file1Byte != -1));

            }

            // if the last read bytes are equal, all bytes are equal
            return file1Byte == file2Byte;
        }
    }
}