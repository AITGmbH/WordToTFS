using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using AIT.TFS.SyncService.Contracts.ProgressService;

namespace AIT.TFS.SyncService.Service.Progress
{
    /// <summary>
    /// Implementation of progress service. See <see cref="IProgressService"/>.
    /// </summary>
    internal class ProgressService : IProgressService
    {

        #region Fields
        private readonly List<ProgressGroup> _groups = new List<ProgressGroup>();
        private IProgressBridgeModel _model;
        private string _title;
        #endregion Fields

        #region Implementation of IProgressService interface
        /// <summary>
        /// Method attaches a progress model.
        /// Model contains relevant properties to describe the progress status and progress status is showed in other part of the application.
        /// </summary>
        /// <param name="model">Model to attach. All relevant data of progress status will be set into this model.</param>
        public void AttachModel(IProgressBridgeModel model)
        {
            if (model == null)
            {
                throw new ArgumentNullException("model");
            }

            _model = model;
            SetProgressTitle();
            SetProgressText();
            SetProgressValue();
        }

        /// <summary>
        /// Method detaches a progress model. No error occurred if no model attached.
        /// </summary>
        public void DetachModel()
        {
            if (_model != null)
            {
                _model.ProgressTitle = string.Empty;
                _model.ProgressValue = 0;
                _model.ProgressText = string.Empty;
            }

            _model = null;
        }

        /// <summary>
        /// Creates new progress process. Any actual settings will be discarded.
        /// </summary>
        /// <param name="progressTitle">Text to use as progress title.</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public void NewProgress(string progressTitle)
        {
            _title = progressTitle;
            if (_model != null)
            {
                _model.ProgressTitle = _title;
            }

            _groups.Clear();
        }

        /// <summary>
        /// Enter the new progress group.
        /// In actual tick of actual progress group will be new progress group created. New progress group contains specified count of ticks.
        /// </summary>
        /// <param name="tickCount">Count of ticks in new progress group.</param>
        /// <remarks>
        /// Progress text will be copied from previous progress group.
        /// </remarks>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public void EnterProgressGroup(int tickCount)
        {
            string statusText = string.Empty;
            if (_groups.Count > 0)
            {
                statusText = _groups[_groups.Count - 1].Text;
            }

            EnterProgressGroup(tickCount, statusText);
        }

        /// <summary>
        /// Enter the new progress group.
        /// In actual tick of actual progress group will be new progress group created. New progress group contains specified count of ticks.
        /// </summary>
        /// <param name="tickCount">Count of ticks in new progress group.</param>
        /// <param name="progressText">Status text of progress.</param>
        /// <remarks>
        /// This is one of two methods, in which status text of progress can be changed. See <see cref="DoTick()"/> and <see cref="DoTick(string)"/>.
        /// </remarks>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public void EnterProgressGroup(int tickCount, string progressText)
        {
            var newGroup = new ProgressGroup();
            newGroup.CountOfTicks = tickCount;
            newGroup.ActualTick = 0;
            newGroup.Text = progressText;
            _groups.Add(newGroup);
            SetProgressText();
        }

        /// <summary>
        /// Leave the actual progress group.
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public void LeaveProgressGroup()
        {
            if (_groups.Count > 0)
            {
                _groups.RemoveAt(_groups.Count - 1);
            }

            SetProgressText();
        }

        /// <summary>
        /// Do one tick in actual progress group. Status text of progress is unchanged.
        /// </summary>
        /// <remarks>
        /// This is one of two methods in which value of progress can be changed. See <see cref="DoTick(string)"/>.
        /// </remarks>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public void DoTick()
        {
            if (_groups.Count <= 0)
            {
                return;
            }

            ProgressGroup pr = _groups[_groups.Count - 1];
            if (pr.ActualTick < pr.CountOfTicks)
            {
                pr.ActualTick += 1;
            }

            SetProgressValue();
        }

        /// <summary>
        /// Do one tick in actual progress group and set new status text of progress.
        /// </summary>
        /// <param name="progressText">New status text of progress.</param>
        /// <remarks>
        /// <para>
        /// This is one of two methods in which value of progress can be changed. See <see cref="DoTick()"/>.
        /// </para><para>
        /// This is one of two methods, in which status text of progress can be changed. See <see cref="EnterProgressGroup(int, string)"/>.
        /// </para>
        /// </remarks>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public void DoTick(string progressText)
        {
            if (_groups.Count <= 0)
            {
                return;
            }

            _groups[_groups.Count - 1].Text = progressText;
            SetProgressText();
            DoTick();
        }

        /// <summary>
        /// Gets the information if the long operation is to cancel.
        /// </summary>
        /// <returns>True for cancel the progress, otherwise false.</returns>
        public bool ProgressCanceled
        {
            get
            {
                if (_model != null)
                {
                    return _model.ProgressCanceled;
                }

                return false;
            }
        }
        /// <summary>
        /// Show progress window.
        /// </summary>
        public void ShowProgress()
        {
            var showHideDialogChanged = ShowHideDialogChanged;
            if (showHideDialogChanged != null)
            {
                showHideDialogChanged(this, new EventArgs());
            }
        }
        /// <summary>
        /// Hide progress window.
        /// </summary>
        public void HideProgress()
        {
            var showHideDialogChanged = ShowHideDialogChanged;
            if (showHideDialogChanged != null)
            {
                showHideDialogChanged(this, new EventArgs());
            }
        }
        /// <summary>
        /// Gets whether progress window is visible.
        /// </summary>
        public bool IsVisibleProgressWindow { get { return _model != null; } }

        /// <summary>
        /// Event called when dialog is going to show or hide.
        /// </summary>
        public event EventHandler ShowHideDialogChanged;

        /// <summary>
        /// Gets all count ticks for actual(last) progress group.
        /// </summary>
        public int ActualProgressGroupCountOfTicks
        {
            get
            {
                if (_groups.Count > 0)
                {
                    return _groups.Last().CountOfTicks;
                }

                return 0;
            }
        }
        /// <summary>
        /// Gets actual count tick for actual(last) progress group.
        /// </summary>
        public int ActualProgressGroupActualTick
        {
            get
            {
                if (_groups.Count > 0)
                {
                    var actualTick = _groups.Last().ActualTick;
                    return actualTick == ActualProgressGroupCountOfTicks ? ActualProgressGroupCountOfTicks : actualTick + 1;
                }

                return 0;
            }
        }

        #endregion Implementation of IProgressService interface

        #region Private methods

        /// <summary>
        /// Method sets the progress title in attached model.
        /// </summary>
        private void SetProgressTitle()
        {
            if (_model != null)
            {
                _model.ProgressTitle = _title;
            }
        }

        /// <summary>
        /// Method sets the progress value in attached model.
        /// </summary>
        private void SetProgressValue()
        {
            if (_model != null)
            {
                double value = 0;
                double previousPart = 1;
                foreach (ProgressGroup pr in _groups)
                {
                    value += previousPart * pr.ActualTick / pr.CountOfTicks;
                    previousPart = previousPart / pr.CountOfTicks;
                }

                _model.ProgressValue = (int) (value * 100);
            }
        }

        /// <summary>
        /// Method sets the progress text in attached model.
        /// </summary>
        private void SetProgressText()
        {
            string text = string.Empty;
            if (_groups.Count > 0)
            {
                text = _groups[_groups.Count - 1].Text;
            }

            if (_model != null)
            {
                _model.ProgressText = text;
            }
        }

        #endregion Private methods
    }
}