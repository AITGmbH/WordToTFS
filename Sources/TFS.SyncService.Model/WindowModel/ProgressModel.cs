#region Usings
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Model.WindowModelBase;
#endregion
namespace AIT.TFS.SyncService.Model.WindowModel
{
    /// <summary>
    /// This class implements a model for window binding to show a progress of long running processes.
    /// </summary>
    /// <remarks>
    /// <para>Class used as a bridge between GUI (progress window) and progress service.</para>
    /// Use the <see cref="Value"/> and <see cref="Text"/> property for binding to show the status of progress.
    /// </remarks>
    public sealed class ProgressModel : ExtBaseModel, IProgressBridgeModel
    {
        #region Private fields
        private string _progressText = string.Empty;
        private string _progressTitle = string.Empty;
        private int _progressValue;
        #endregion
        #region Public properties
        /// <summary>
        /// Gets the actual title of progress.
        /// </summary>
        /// <remarks>Binds this property in GUI window.</remarks>
        public string Title
        {
            get { return _progressTitle; }
            set
            {
                if (_progressTitle != value)
                {
                    _progressTitle = value;
                    TriggerPropertyChanged(this, "ProgressTitle");
                    TriggerPropertyChanged(this, nameof(Title));
                }
            }
        }

        /// <summary>
        /// Gets the actual value of progress. Value should be in interval &lt;0, 100&gt;.
        /// </summary>
        /// <remarks>Bind this property in GUI window.</remarks>
        public int Value
        {
            get { return _progressValue; }
            set
            {
                if (_progressValue != value)
                {
                    _progressValue = value;
                    TriggerPropertyChanged(this, "ProgressValue");
                    TriggerPropertyChanged(this, nameof(Value));
                }
            }
        }

        /// <summary>
        /// Gets the actual text of progress.
        /// </summary>
        /// <remarks>Bind this property in GUI window.</remarks>
        public string Text
        {
            get { return _progressText; }
            set
            {
                if (_progressText != value)
                {
                    _progressText = value;
                    TriggerPropertyChanged(this, "ProgressText");
                    TriggerPropertyChanged(this, nameof(Text));
                }
            }
        }
        #endregion
        #region IProgressBridgeModel Members

        /// <summary>
        /// Gets or sets the title of the whole progress process.
        /// </summary>
        string IProgressBridgeModel.ProgressTitle
        {
            get { return _progressTitle; }
            set
            {
                if (_progressTitle != value)
                {
                    _progressTitle = value;
                    TriggerPropertyChanged(this, "ProgressTitle");
                    TriggerPropertyChanged(this, nameof(Title));
                }
            }
        }

        /// <summary>
        /// Gets or sets the value of progress. Value is from interval &lt;0, 100&gt;.
        /// </summary>
        int IProgressBridgeModel.ProgressValue
        {
            get { return _progressValue; }
            set
            {
                if (_progressValue != value)
                {
                    _progressValue = value;
                    TriggerPropertyChanged(this, "ProgressValue");
                    TriggerPropertyChanged(this, nameof(Value));
                }
            }
        }

        /// <summary>
        /// Gets or sets the text of progress.
        /// </summary>
        string IProgressBridgeModel.ProgressText
        {
            get { return _progressText; }
            set
            {
                if (_progressText != value)
                {
                    _progressText = value;
                    TriggerPropertyChanged(this, "ProgressText");
                    TriggerPropertyChanged(this, nameof(Text));
                }
            }
        }

        /// <summary>
        /// Gets the information if the long operation is to cancel.
        /// </summary>
        /// <returns>True for cancel the progress, otherwise false.</returns>
        public bool ProgressCanceled { get; set; }

        #endregion
    }
}