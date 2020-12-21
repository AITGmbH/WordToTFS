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
    public class ComInteropException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        protected ComInteropException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        public ComInteropException()
            :base(Resources.MappingExceptionMessage)
        {
            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        public ComInteropException(string message)
            : base(message)
        {
            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The inner exception.</param>
        public ComInteropException(string message, Exception innerException)
            :base(message,innerException)
        {
            
        }
    }
}
