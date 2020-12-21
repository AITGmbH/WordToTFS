using System;
using AIT.TFS.SyncService.Contracts.Properties;
using System.Runtime.Serialization;

namespace AIT.TFS.SyncService.Contracts.Exceptions
{
    /// <summary>
    /// Use this class for all exceptions that can occur during the work with the interfaces of word and office interop.
    /// 
    /// </summary>
    [Serializable]
    public class AttachmentException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        protected AttachmentException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        public AttachmentException()
            :base(Resources.AttachmentExceptionMessage)
        {
            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        public AttachmentException(string message)
            : base(message)
        {
            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The inner exception.</param>
        public AttachmentException(string message, Exception innerException)
            :base(message,innerException)
        {
            
        }
    }
}
