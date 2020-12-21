using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace AIT.TFS.SyncService.Contracts.Exceptions
{
    /// <summary>
    /// Class containing the IOleMessageFilter thread error-handling functions.
    /// </summary>
    [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Reviewed. Suppression is OK here.")]
    public class MessageFilter : IOleMessageFilter
    {
        // ReSharper disable InconsistentNaming
        private const int S_OK = 0;
        private const int SERVERCALL_ISHANDLED = 0;
        private const int SERVERCALL_RETRYLATER = 2;
        private const int PENDINGMSG_WAITDEFPROCESS = 2;
        // ReSharper restore InconsistentNaming

        /// <summary>
        /// Apply this filter.
        /// </summary>
        public static bool Register()
        {
            IOleMessageFilter newFilter = new MessageFilter();
            IOleMessageFilter oldFilter;
            var result = NativeMethods.CoRegisterMessageFilter(newFilter, out oldFilter);
            return result == S_OK;
        }

        /// <summary>
        /// Done with the filter, close it.
        /// </summary>
        public static bool Revoke()
        {
            IOleMessageFilter oldFilter;
            var result = NativeMethods.CoRegisterMessageFilter(null, out oldFilter);
            return result == S_OK;
        }

        /// <summary>
        /// Handles an incoming call
        /// </summary>
        public int HandleIncomingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            // There are no incoming calls anyways.
            return SERVERCALL_ISHANDLED;
        }

        /// <summary>
        /// Called when a COM call was rejected.
        /// </summary>
        public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == SERVERCALL_RETRYLATER)
            {
                // Waittime in milliseconds. A value < 100 means retry immediately
                return 100;
            }

            // Too busy; cancel call.
            return -1;
        }

        /// <summary>
        /// Window Message arrived before COM call was completed.
        /// </summary>
        public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            // Just wait.
            return PENDINGMSG_WAITDEFPROCESS;
        }
    }

    /// <summary>
    /// Native methods used by the message filter class.
    /// </summary>
    internal static class NativeMethods
    {
        [DllImport("Ole32.dll")]
        internal static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    }
}