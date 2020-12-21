using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using AIT.TFS.SyncService.Contracts;
using System.Diagnostics.CodeAnalysis;

namespace AIT.TFS.SyncService.Factory
{
    /// <summary>
    /// Class implements the functionality of a tracer service.
    /// </summary>
    public static class SyncServiceTrace
    {
        private const string Exception = "Exception";
        private const string Information = "Information";
        private const string Warning = "Warning";
        private const string Debug = "Debug";

        private const string LogFileName = "logs.txt";

        /// <summary>
        /// Main application debug trace switch. Change this to define how verbose the logging output should be
        /// </summary>
        public static TraceSwitch DebugLevel = new TraceSwitch("DebugLevel", "The Output level of tracing", "0");

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncServiceTrace"/> class.
        /// </summary>
        static SyncServiceTrace()
        {
            Trace.AutoFlush = true;
            Trace.Listeners.Add(new TextWriterTraceListener(LogFile, "WordToTFSTraceListener"));
        }

        /// <summary>
        /// Gets the path to the log file.
        /// </summary>
        public static string LogFile
        {
            get
            {
                string directory =
                    Path.Combine(
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                     Constants.ApplicationCompany),
                        Constants.ApplicationName);

                Directory.CreateDirectory(directory);

                return Path.Combine(directory, LogFileName);
            }
        }

        
        /// <summary>
        /// Method logs the exception.
        /// </summary>
        /// <param name="ex">Exception to log.</param>
        public static void LogException(Exception ex)
        {
            if (ex != null)
            {
                Write(Exception, $"M: {ex}");
            }
        }

        /// <summary>
        /// Logs an information. Usage like string.Format
        /// </summary>
        /// <param name="information">Information message with format items.</param>
        /// <param name="parameter">Objects that replace the corresponding format items</param>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
        public static void I(string information, params object[] parameter)
        {
            Write(Information, string.Format(CultureInfo.InvariantCulture, information, parameter));
        }

        /// <summary>
        /// Logs an debug information. These are only visible if tracing is set to verbose. Usage like string.Format
        /// </summary>
        /// <param name="debug">Debug message with format items.</param>
        /// <param name="parameter">Objects that replace the corresponding format items</param>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
        public static void D(string debug, params object[] parameter)
        {
            Write(Debug, string.Format(CultureInfo.InvariantCulture, debug, parameter));
        }

        /// <summary>
        /// Logs an warning. These are only visible if tracing is set to warning or above. Usage like string.Format
        /// </summary>
        /// <param name="warning">Debug message with format items.</param>
        /// <param name="parameter">Objects that replace the corresponding format items</param>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
        public static void W(string warning, params object[] parameter)
        {
            Write(Warning, string.Format(CultureInfo.InvariantCulture, warning, parameter));
        }

        /// <summary>
        /// Logs an exception. Exceptions are always visible unless logging is turned off. Usage like string.Format
        /// </summary>
        /// <param name="exception">Debug message with format items.</param>
        /// <param name="parameter">Objects that replace the corresponding format items</param>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
        public static void E(string exception, params object[] parameter)
        {
            Write(Exception, string.Format(CultureInfo.InvariantCulture, exception, parameter));
        }

        /// <summary>
        /// Method writes the message to log.
        /// </summary>
        /// <param name="category">Category of the message to log.</param>
        /// <param name="message">Message to log.</param>
        private static void Write(string category, string message)
        {
            if (category == Warning)
            {
                Trace.WriteLineIf(DebugLevel.TraceWarning, FormatMessage(message, category));
            }
            else if (category == Exception)
            {
                Trace.WriteLineIf(DebugLevel.TraceError, FormatMessage(message, category));
            }
            else if (category == Information)
            {
                Trace.WriteLineIf(DebugLevel.TraceInfo, FormatMessage(message, category));
            }
            else if (category == Debug)
            {
                Trace.WriteLineIf(DebugLevel.TraceVerbose, FormatMessage(message, category));
            }
        }

        /// <summary>
        /// Adds a time stamp to a <paramref name="logEntry"/>.
        /// </summary>
        /// <param name="logEntry">Raw message.</param>
        /// <param name="category">Log detail (Error, Warning, Information, Debug)</param>
        /// <returns>Formatted log entry.</returns>
        private static string FormatMessage(string logEntry, string category)
        {
            return string.Format(CultureInfo.CurrentCulture, "{0} ({1}) {2}", DateTime.Now, category.PadRight(11, ' '), logEntry);
        }
    }
}