#region Usings
using System;
using System.IO;
using System.Reflection;
#endregion

namespace AIT.TFS.SyncService.Model.Console
{
    /// <summary>
    /// Helper to write all information to the console and to a debug log
    /// </summary>
    public static class ConsoleExtensionLogging
    {
        private static string fileName;

        /// <summary>
        /// Specifies the log level of the console.
        /// File: Log only to file
        /// Console: Log to system console
        /// Both: Log to both
        /// </summary>
        public enum LogLevel
        {
            Both, Console, LogFile
        };

        /// <summary>
        /// Logs a message to the target specified with the enum LogLevel
        /// </summary>
        /// <param name="message"></param>
        /// <param name="level"></param>
        public static void LogMessage(string message, LogLevel level)
        {
            switch (level)
            {

                case LogLevel.Console:
                    System.Console.WriteLine(message);
                    break;
                case LogLevel.LogFile:
                    WriteMessageToFile(message);
                    break;
                case LogLevel.Both:
                    System.Console.WriteLine(message);
                    WriteMessageToFile(message);
                    break;
            }
        }


        /// <summary>
        /// Write a message to a given file
        /// </summary>
        /// <param name="message"></param>
        private static void WriteMessageToFile(string message)
        {
            if (String.IsNullOrEmpty(fileName))
            {
                fileName = DateTime.Now.ToString(@"yyyy_MM_dd") + ".log";

            }
            var executionPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            try
            {
                using (StreamWriter w = File.AppendText(executionPath + "\\" + fileName))
                {
                    AppendLog(message, w);
                }
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Append log message to a given text writer
        /// </summary>
        /// <param name="logMessage"></param>
        /// <param name="txtWriter"></param>
        private static void AppendLog(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.Write("\r\nLog Entry : ");
                txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                    DateTime.Now.ToLongDateString());
                txtWriter.WriteLine("  :{0}", logMessage);
                txtWriter.WriteLine("-------------------------------");
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {
            }
        }
    }
}
