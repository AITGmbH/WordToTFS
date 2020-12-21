using System;

namespace AIT.TFS.SyncService.Factory
{
    /// <summary>
    /// Convenience class for argument checkers
    /// </summary>
    public static class Guard
    {
        /// <summary>
        /// Helper attribute that marks a variable as checked against null.
        /// </summary>
        private sealed class ValidatedNotNullAttribute : Attribute
        {
        }

        /// <summary>
        /// Throws exception if argument is null
        /// </summary>
        /// <param name="reference">The reference to check against null.</param>
        /// <param name="argumentName">Name of the argument.</param>
        /// <exception cref="System.ArgumentNullException">Thrown if reference is null.</exception>
        public static void ThrowOnArgumentNull<T>([ValidatedNotNull] T reference, string argumentName) where T : class
        {
            if (reference == null)
            {
                throw new ArgumentNullException(argumentName);
            }
        }

        /// <summary>
        /// Throws exception if argument is null
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="reference">The reference to check against null.</param>
        /// <param name="argumentName">Name of the argument.</param>
        /// <param name="message">Optional message.</param>
        /// <exception cref="System.ArgumentNullException">Thrown if reference is null.</exception>
        public static void ThrowOnArgumentNull<T>([ValidatedNotNull] T reference, string argumentName, string message) where T : class
        {
            if (reference == null)
            {
                throw new ArgumentNullException(argumentName, message);
            }
        }
    }
}
