using System;
namespace AIT.TFS.SyncService.Contracts.ProgressService
{
    /// <summary>
    /// Interface defines functionality of progress service.
    /// </summary>
    public interface IProgressService
    {
        /// <summary>
        /// Gets the information if the long operation is to cancel.
        /// </summary>
        /// <returns>True for cancel the progress, otherwise false.</returns>
        bool ProgressCanceled { get; }

        /// <summary>
        /// Method attaches a progress model.
        /// Model contains relevant properties to describe the progress status and progress status is showed in other part of the application.
        /// </summary>
        /// <param name="model">Model to attach. All relevant data of progress status will be set into this model.</param>
        void AttachModel(IProgressBridgeModel model);

        /// <summary>
        /// Method detaches a progress model. No error occurred if no model attached.
        /// </summary>
        void DetachModel();

        /// <summary>
        /// Creates new progress process. Any actual settings will be discarded.
        /// </summary>
        /// <param name="progressTitle">Text to use as progress title.</param>
        void NewProgress(string progressTitle);

        /// <summary>
        /// Enter the new progress group.
        /// In actual tick of actual progress group will be new progress group created. New progress group contains specified count of ticks.
        /// </summary>
        /// <param name="tickCount">Count of ticks in new progress group.</param>
        /// <remarks>
        /// Progress text will be copied from previous progress group.
        /// </remarks>
        void EnterProgressGroup(int tickCount);

        /// <summary>
        /// Enter the new progress group.
        /// In actual tick of actual progress group will be new progress group created. New progress group contains specified count of ticks.
        /// </summary>
        /// <param name="tickCount">Count of ticks in new progress group.</param>
        /// <param name="progressText">Status text of progress.</param>
        /// <remarks>
        /// This is one of two methods, in which status text of progress can be changed. See <see cref="DoTick()"/> and <see cref="DoTick(string)"/>.
        /// </remarks>
        void EnterProgressGroup(int tickCount, string progressText);

        /// <summary>
        /// Leave the actual progress group.
        /// </summary>
        void LeaveProgressGroup();

        /// <summary>
        /// Do one tick in actual progress group. Status text of progress is unchanged.
        /// </summary>
        /// <remarks>
        /// This is one of two methods in which value of progress can be changed. See <see cref="DoTick(string)"/>.
        /// </remarks>
        void DoTick();

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
        void DoTick(string progressText);
        /// <summary>
        /// Show progress window.
        /// </summary>
        void ShowProgress();
        /// <summary>
        /// Hide progress window.
        /// </summary>
        void HideProgress();
        /// <summary>
        /// Gets whether progress window is visible.
        /// </summary>
        bool IsVisibleProgressWindow { get; }
        /// <summary>
        /// Event called when dialog is going to show or hide.
        /// </summary>
        event EventHandler ShowHideDialogChanged;
         /// <summary>
        /// Gets all count ticks for actual(last) progress group.
        /// </summary>
        int ActualProgressGroupCountOfTicks { get; }
        /// <summary>
        /// Gets actual count tick for actual(last) progress group.
        /// </summary>
        int ActualProgressGroupActualTick { get; }
    }
}