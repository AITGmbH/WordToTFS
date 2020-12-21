using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Service;
using TFS.SyncService.View.Word2007.Controls;
using TFS.SyncService.View.Word2007.Properties;
using Office = Microsoft.Office.Core;

namespace TFS.SyncService.View.Word2007
{
    using System;
    using System.Collections.ObjectModel;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;
    using System.Windows.Forms;

    using AIT.TFS.SyncService.Common.Helper;
    using AIT.TFS.SyncService.Factory;
    using AIT.TFS.SyncService.Model;

    using Microsoft.Office.Interop.Word;
    using Microsoft.Office.Tools;
    using AIT.TFS.SyncService.Model.WindowModel;


    [ComVisible(true)]
    public partial class ThisAddIn
    {
        #region Private fields

        /// <summary>
        /// Holds WordToTFS ribbon.
        /// </summary>
        private WordRibbon _ribbon;

        #endregion Private fields

        #region Public properties

        /// <summary>
        /// Gets custom task panes without orphaned task panes.
        /// </summary>
        public CustomTaskPaneCollection WordToTfsTaskPanes
        {
            get
            {
                RemoveOrphanedTaskPanes();
                return CustomTaskPanes;
            }
        }

        #endregion Public properties

        #region Public methods

        /// <summary>
        /// Initialize the word AddIn
        /// </summary>

        [ComVisible(true)]
        public void InitializeAddIn()
        {
            // Find the ribbon for word2tfs
            foreach (var ribbon in Globals.Ribbons)
            {
                if (ribbon is WordRibbon)
                {
                    _ribbon = (WordRibbon)ribbon;
                }
            }

            try
            {
                var serviceModel = new SyncServiceModel(Application);
                if (_ribbon != null)
                {
                    _ribbon.AttachModel(serviceModel);
                }

                (Application as ApplicationEvents4_Event).NewDocument += x => RemoveOrphanedTaskPanes();

                Application.DocumentOpen += x => RemoveOrphanedTaskPanes();
                Application.DocumentChange += RemoveOrphanedTaskPanes;
            }
            catch (Exception ex)
            {
                SyncServiceTrace.LogException(ex);
            }
        }

        /// <summary>
        /// Gets the information if the task pane for the document is visible.
        /// </summary>
        public bool TaskPaneVisible(_Document document, string caption)
        {
            return CustomTaskPanes.Any(pane => pane.Window == document.ActiveWindow && pane.Title.Equals(caption) && pane.Visible);
        }

        /// <summary>
        /// Gets the information if the task pane for the document is visible.
        /// </summary>
        public bool TaskPaneVisible<T>(_Document document) where T : UIElement
        {
            return CustomTaskPanes.Any(pane => pane.Window == document.ActiveWindow && pane.Control is HostPane<T> && pane.Visible);
        }

        /// <summary>
        /// Shows the task pane with the specified content.
        /// Either turns an existing pane visible or creates a new task pane
        /// </summary>
        public void ShowTaskPane(_Document document, UserControl content, string title)
        {
            if (document == null || content == null)
            {
                return;
            }

            var pane = FindTaskPane(document, content);
            if (pane == null)
            {
                pane = CustomTaskPanes.Add(content, title, document.ActiveWindow);
                pane.VisibleChanged += HandleCustomTaskPaneVisibleChanged;
            }

            pane.Visible = true;
        }

        /// <summary>
        /// Method hides the task pane for the document.
        /// </summary>
        /// <param name="document">Word document to hide task pane for.</param>
        /// <param name="content">The content of the task pane to hide</param>
        public void HideTaskPane(_Document document, UserControl content)
        {
            var pane = FindTaskPane(document, content);
            if (pane != null)
            {
                pane.Visible = false;
            }
        }

        #endregion Public methods

        #region Private methods

        /// <summary>
        /// Removes orphaned task panes. When gets a task pane orphaned? - SER
        /// </summary>
        private void RemoveOrphanedTaskPanes()
        {
            var ctpc = new Collection<CustomTaskPane>();

            foreach (var ctp in CustomTaskPanes)
            {
                try
                {
                    if (ctp.Window == null)
                    {
                        ctpc.Add(ctp);
                    }
                }
                catch (COMException)
                {
                    // "Task Pane no longer valid" ...
                    ctpc.Add(ctp);
                }
            }

            foreach (var ctp in ctpc)
            {
                CustomTaskPanes.Remove(ctp);
            }
        }

        private CustomTaskPane FindTaskPane(_Document document, UserControl content)
        {
            if (document == null)
            {
                return null;
            }
            return CustomTaskPanes.FirstOrDefault(pane => (pane.Control == content) && (pane.Window == document.ActiveWindow));
        }

        private void HandleCustomTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            var taskPane = sender as CustomTaskPane;

            if (taskPane != null)
            {
                _ribbon.VisibleChanged(taskPane.Control, taskPane.Visible);
            }
        }

        /// <summary>
        /// Called when the add-in is loaded, after all the initialization code in the assembly has run.
        /// </summary>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            //TraceLevel level = (TraceLevel)Settings.Default.EnableDebugging;
            //SyncServiceTrace.DebugLevel.Level = level;
            //SettingsModel.Level = level;
            SyncServiceTrace.I(Resources.LogService_Startup);

            // Register custom message filter for COM calls. See http://msdn.microsoft.com/en-us/library/ms228772%28v=vs.90%29.aspx
            if (!MessageFilter.Register())
            {
                SyncServiceTrace.W("Failed to register message handler.");
            }

            // Initialize all assemblies.
            AssemblyInit.Instance.Init();
            AIT.TFS.SyncService.Adapter.TFS2012.AssemblyInit.Instance.Init();
            AIT.TFS.SyncService.Adapter.Word2007.AssemblyInit.Instance.Init();

            InitializeAddIn();

            // if you enable / diable the addin, _ribbon is not set
            if (_ribbon != null)
            {
                _ribbon.RegisterToServices();
            }

            SyncServiceTrace.I(Resources.LogService_StartupComplete);
        }

        /// <summary>
        /// Called when the add-in is about to be unloaded.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (!MessageFilter.Revoke())
            {
                SyncServiceTrace.W("Failed to unregister message handler");
            }
            Settings.Default.Save();
            AssemblyInit.Instance.Dispose();
            AIT.TFS.SyncService.Adapter.TFS2012.AssemblyInit.Instance.Dispose();
            AIT.TFS.SyncService.Adapter.Word2007.AssemblyInit.Instance.Dispose();

            // Make sure the temp folder is not deleted by "Nested Adapters" when generating test reports.
            var counter = Process.GetProcesses().Count(process => process.ProcessName.Contains("WINWORD"));
            if (counter == 1)
            {
                TempFolder.ClearTempFolder();
            }
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;

            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;

            if (exception != null)
            {
                try
                {
                    SyncServiceTrace.E("Unhandled Exception (terminating {0}) {1}", e.IsTerminating, exception);
                }
                catch (Exception ex)
                {
                    SyncServiceTrace.LogException(ex);
                }

                throw exception;
            }
        }

        #endregion Private methods
    }
}
