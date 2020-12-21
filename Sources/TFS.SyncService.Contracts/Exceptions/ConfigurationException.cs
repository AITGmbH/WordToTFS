using System;
using AIT.TFS.SyncService.Contracts.Properties;
using System.Runtime.Serialization;

namespace AIT.TFS.SyncService.Contracts.Exceptions
{
    /// <summary>
    /// Use this class for all exceptions related to errors in the configuration files
    /// The user will most likely want to see an informative message when his w2t configurations
    /// don't work as expected
    /// </summary>
    [Serializable]
    public class ConfigurationException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        protected ConfigurationException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        public ConfigurationException()
            :base(Resources.MappingExceptionMessage)
        {
            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        public ConfigurationException(string message)
            : base(message)
        {
            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The inner exception.</param>
        public ConfigurationException(string message, Exception innerException)
            :base(message,innerException)
        {
            
        }
    }
}
