#region Usings
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Model.Properties;
using AIT.TFS.SyncService.Model.WindowModelBase;
using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;
#endregion

namespace AIT.TFS.SyncService.Model.TemplateManagement
{
    /// <summary>
    /// The class implements model for the <see cref="TemplateManagerModel"/>.
    /// </summary>
    public class TemplateManagerModel : ExtBaseModel
    {
        #region Private fields

        /// <summary>
        /// Holds the show name of new template bundle.
        /// </summary>
        private string _showName;

        /// <summary>
        /// Holds the location of new template bundle.
        /// </summary>
        private string _location;

        /// <summary>
        /// Holds a value indicating whether the operation is active.
        /// </summary>
        private bool _operationActive;

        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateManagerModel"/> class.
        /// </summary>
        /// <param name="templateManager">Template manager.</param>
        public TemplateManagerModel(TemplateManager templateManager)
        {
            TemplateManager = templateManager;
            SetTemplateBundleLocation = new ViewCommand(ExecuteSetTemplateBundleLocation);
            AddTemplateBundle = new ViewCommand(ExecuteAddTemplateBundle, CanExecuteAddTemplateBundle);
            RemoveTemplateBundle = new ViewCommand(ExecuteRemoveTemplateBundle, CanExecuteRemoveTemplateBundle);
            ReloadTemplateBundle = new ViewCommand(ExecuteReloadTemplateBundle);
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the template manager.
        /// </summary>
        public TemplateManager TemplateManager
        {
            get;
            private set;
        }

        /// <summary>
        /// <see cref="ICommand"/> to set the template bundle location.
        /// </summary>
        public ViewCommand SetTemplateBundleLocation
        {
            get;
            private set;
        }

        /// <summary>
        /// <see cref="ICommand"/> to add the template bundle.
        /// </summary>
        public ViewCommand AddTemplateBundle
        {
            get;
            private set;
        }

        /// <summary>
        /// <see cref="ICommand"/> to remove the template bundle.
        /// </summary>
        public ViewCommand RemoveTemplateBundle
        {
            get;
            private set;
        }

        /// <summary>
        /// <see cref="ICommand"/> to reload the template bundle.
        /// </summary>
        public ViewCommand ReloadTemplateBundle
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets the show name of new template bundle.
        /// </summary>
        public string ShowName
        {
            get
            {
                return _showName;
            }
            set
            {
                if (_showName != value)
                {
                    _showName = value;
                    TriggerPropertyChanged(nameof(ShowName));
                    AddTemplateBundle.CallEventCanExecuteChanged();
                }
            }
        }

        /// <summary>
        /// Gets or sets the location of new template bundle.
        /// </summary>
        public string Location
        {
            get
            {
                return _location;
            }
            set
            {
                if (_location != value)
                {
                    _location = value;
                    TriggerPropertyChanged(nameof(Location));
                    AddTemplateBundle.CallEventCanExecuteChanged();
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the operation is active
        /// </summary>
        public bool OperationActive
        {
            get
            {
                return _operationActive;
            }
            set
            {
                if (_operationActive != value)
                {
                    _operationActive = value;
                    TriggerPropertyChanged(nameof(OperationActive));
                }
            }
        }

        #endregion Public properties

        #region Private command methods

        /// <summary>
        /// 'Execute' method of <see cref="SetTemplateBundleLocation"/> command.
        /// </summary>
        /// <param name="parameter">Command parameter.</param>
        public void ExecuteSetTemplateBundleLocation(object parameter)
        {
            using (var browserDialog = new FolderBrowserDialog())
            {
                browserDialog.Description = Resources.BrowserDialogDescription;
                if (browserDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                Location = browserDialog.SelectedPath;
            }
        }

        /// <summary>
        /// 'Execute' method of <see cref="AddTemplateBundle"/> command.
        /// </summary>
        /// <param name="parameter">Command parameter.</param>
        public void ExecuteAddTemplateBundle(object parameter)
        {
            var loadAsProjectMappedFolderHierarchy = false;

            if (Directory.Exists(Location) && Directory.GetDirectories(Location).Length > 0)
            {
                var dialogResult = MessageBox.Show(Resources.TM_AddTemplateBundleAsProjectMappedHierarchy, Resources.MessageBox_Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, 0);
                loadAsProjectMappedFolderHierarchy = dialogResult == DialogResult.Yes;
            }

            OperationActive = true;
            var backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += (s, e) => e.Result = new TemplateBundle(ShowName, Location, Guid.NewGuid());
            backgroundWorker.RunWorkerCompleted += (s, e) =>
            {
                if (e.Error != null)
                {
                    SyncServiceTrace.LogException(e.Error);
                }
                else if (e.Result != null)
                {
                    var templateBundle = (TemplateBundle)e.Result;
                    templateBundle.HasProjectMappedFolderHierarchy = loadAsProjectMappedFolderHierarchy;
                    TemplateManager.AddTemplateBundle((TemplateBundle)e.Result);
                }
                OperationActive = false;
                backgroundWorker.Dispose();
            };
            backgroundWorker.RunWorkerAsync();
        }

        /// <summary>
        /// 'CanExecute' method of <see cref="AddTemplateBundle"/> command.
        /// </summary>
        /// <param name="parameter">Command parameter.</param>
        /// <returns><c>true</c> to enable command, otherwise <c>false</c>.</returns>
        public bool CanExecuteAddTemplateBundle(object parameter)
        {
            return !string.IsNullOrEmpty(ShowName) && !string.IsNullOrEmpty(Location);
        }

        /// <summary>
        /// 'Execute' method of <see cref="RemoveTemplateBundle"/> command.
        /// </summary>
        /// <param name="parameter">Command parameter.</param>
        public void ExecuteRemoveTemplateBundle(object parameter)
        {
            // Parameter is the template bundle.
            var templateBundle = (string)parameter;

            // Ask user if he is really sure about deleting the bundle
            var result = MessageBox.Show(string.Format(CultureInfo.CurrentCulture, Resources.TM_ConfirmDeleteText, templateBundle), Resources.TM_ConfirmDeleteHeader, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2, 0);
            if (result != DialogResult.Yes)
            {
                return;
            }

            // TODO: WTF?? Is this a non-blocking version of Thread.Sleep(500)? Why not use a timer and why do I have to sleep in the first place? - SER
            OperationActive = true;
            var bw = new BackgroundWorker();
            bw.DoWork += (s, e) => Thread.Sleep(500);
            bw.RunWorkerCompleted += (s, e) =>
            {
                TemplateManager.RemoveTemplateBundle(templateBundle);
                OperationActive = false;
                bw.Dispose();
            };
            bw.RunWorkerAsync();
        }

        /// <summary>
        /// 'Execute' method of <see cref="ReloadTemplateBundle"/> command.
        /// </summary>
        /// <param name="parameter">Name of the template bundle to reload.</param>
        private void ExecuteReloadTemplateBundle(object parameter)
        {
            var templateBundleName = (string)parameter;
            TemplateManager.ReloadTemplateBundle(templateBundleName);
        }

        /// <summary>
        /// 'CanExecute' method of <see cref="RemoveTemplateBundle"/> command.
        /// </summary>
        /// <param name="parameter">Command parameter.</param>
        /// <returns><c>true</c> to enable command, otherwise <c>false</c>.</returns>
        public bool CanExecuteRemoveTemplateBundle(object parameter)
        {
            return true;
        }

        #endregion Private command methods
    }
}
