#region Usings
using System;
using System.Runtime.InteropServices;
#endregion
namespace TFS.SyncService.View.Word2007
{
    /// <summary>
    /// Native methods used the WordToTFS plugin
    /// </summary>
    internal static class NativeMethods
    {
        [DllImport("user32", ExactSpelling = false, CharSet = CharSet.Unicode)]
        internal static extern IntPtr FindWindow(string className, string windowName);
    }
}
