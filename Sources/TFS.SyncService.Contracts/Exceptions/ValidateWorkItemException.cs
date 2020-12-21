using System;
using System.Runtime.Serialization;

namespace AIT.TFS.SyncService.Contracts.Exceptions
{
    /// <summary>
    /// <see cref="ValidateWorkItemException"/> implements exception thrown by generation of test report.
    /// </summary>
    [Serializable]
    public class ValidateWorkItemException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ValidateWorkItemException"/> class.
        /// </summary>
        public ValidateWorkItemException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ValidateWorkItemException"/> class.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public ValidateWorkItemException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ValidateWorkItemException"/> class.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception, or a null reference if no inner exception is specified.</param>
        public ValidateWorkItemException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ValidateWorkItemException"/> class.
        /// </summary>
        /// <param name="info">The <see cref="SerializationInfo"/> that holds the serialized object data about the exception being thrown.</param>
        /// <param name="context">The <see cref="StreamingContext"/> that contains contextual information about the source or destination.</param>
        protected ValidateWorkItemException(SerializationInfo info, StreamingContext context)
            :base(info, context)
        {
        }
    }
}
