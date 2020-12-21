using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace AIT.TFS.SyncService.Contracts.Exceptions
{
    /// <summary>
    /// This interface is used to register custom OLE exception handling.
    /// COM has problems with multithreading which may lead to "Application is busy"-Exceptions.
    /// </summary>
    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleMessageFilter
    {
        /// <summary>
        /// Handles an incoming call
        /// </summary>
        [PreserveSig]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", Justification = "Using Win32 naming for consistency.")]
        [SuppressMessage("Microsoft.StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Using Win32 naming for consistency.")]
        int HandleIncomingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        /// <summary>
        /// Called when a COM call was rejected.
        /// </summary>
        [PreserveSig]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", Justification = "Using Win32 naming for consistency.")]
        [SuppressMessage("Microsoft.StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Using Win32 naming for consistency.")]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

        /// <summary>
        /// Window Message arrived before COM call was completed.
        /// </summary>
        [PreserveSig]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", Justification = "Using Win32 naming for consistency.")]
        [SuppressMessage("Microsoft.StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Using Win32 naming for consistency.")]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }
}