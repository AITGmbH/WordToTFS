using System;

namespace AIT.TFS.SyncService.Contracts.Model
{
    /// <summary>
    /// The class implements event args for <see cref="ISyncServiceModel"/>.
    /// </summary>
    public class SyncServiceEventArgs : EventArgs
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncServiceEventArgs"/> class.
        /// </summary>
        /// <param name="documentName">Name of the document.</param>
        /// <param name="documentFullName">Full name of the document.</param>
        public SyncServiceEventArgs(string documentName, string documentFullName)
        {
            DocumentName = documentName;
            DocumentFullName = documentFullName;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the name of document where is the change of data occurred.
        /// </summary>
        public string DocumentName { get; private set; }

        /// <summary>
        /// Gets the full name of document where is the change of data occurred.
        /// </summary>
        public string DocumentFullName { get; private set; }

        #endregion Public properties
    }
}
