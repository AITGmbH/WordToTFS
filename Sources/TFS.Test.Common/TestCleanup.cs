#region Usings
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using AIT.TFS.SyncService.Factory;
using Microsoft.Office.Interop.Word;
#endregion

namespace TFS.Test.Common
{

    /// <summary>
    /// This class will help to clean the test environment by closing existing documents and killing active word processes
    /// </summary>
    public static class TestCleanup
    {

        /// <summary>
        /// Close all open word instances
        /// </summary>
        public static void CloseWordDocumentAndKillOpenWordInstances()
        {
            SyncServiceTrace.I("Killing word by process");
            foreach (Process p in Process.GetProcessesByName("winword"))
            {
                SyncServiceTrace.I("Found process");
                try
                {
                    SyncServiceTrace.I("Killing process with id:" +p.Id);
                    p.Kill();
                    p.WaitForExit(); // possibly with a timeout
                }
                catch (Win32Exception winException)
                {
                    SyncServiceTrace.I("Exception caught " + winException.Message);
                    // process was terminating or can't be terminated - deal with it
                }
                catch (InvalidOperationException invalidException)
                {
                    SyncServiceTrace.I("Exception caught " + invalidException.Message);
                    // process has already exited - might be able to let this one go
                }
            }

            Boolean openWordApplicationsExist = true;
            while (openWordApplicationsExist)
            {
                try
                {
                    SyncServiceTrace.I("Get word application");

                    var wordApplication = (Application)Marshal.GetActiveObject("Word.Application");

                    if (wordApplication == null)
                    {
                        SyncServiceTrace.I("No word application found");

                        openWordApplicationsExist = false;
                        Thread.CurrentThread.Join(200);
                    }
                    else
                    {
                        SyncServiceTrace.I("Trying to quit word");
                        
                        wordApplication.Quit(false);

                        Thread.CurrentThread.Join(200);
                        SyncServiceTrace.I("Release com");
                        Marshal.FinalReleaseComObject(wordApplication);
                        Thread.CurrentThread.Join(200);

                        SyncServiceTrace.I("Set word application to null");
                        // ReSharper disable once RedundantAssignment
                        wordApplication = null;

                        SyncServiceTrace.I("Calling GC Collect and wait for pending finalizers");
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                       
                    }
                }
                catch (Exception)
                {
                    SyncServiceTrace.I("Exception caught. Cleaning is finished");
                    openWordApplicationsExist = false;
                }
            }

            SyncServiceTrace.I("Finished cleaning");

        }

        /// <summary>
        /// Close a given document
        /// </summary>
        /// <param name="testDocument">Word document</param>
        public static void CloseWordDocument(Document testDocument)
        {
            try
            {
                SyncServiceTrace.I("Close documents");
                testDocument.Close(false);
                SyncServiceTrace.I("Closing successfull");
                SyncServiceTrace.I("Set test document to null");
                // ReSharper disable once RedundantAssignment
                testDocument = null;
                SyncServiceTrace.I("Releasing successfull");
            }
            catch (Exception ex)
            {
                SyncServiceTrace.I("Exception caught " + ex.Message);
            }
        }
    }
}
