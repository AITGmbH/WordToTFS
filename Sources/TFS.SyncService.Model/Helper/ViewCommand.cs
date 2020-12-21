#region Usings
using System;
using System.Windows.Input;
#endregion

namespace AIT.TFS.SyncService.Model.Helper
{
    /// <summary>
    /// The class implements <see cref="ICommand"/> to bind the commands from model in view.
    /// </summary>
    public class ViewCommand : ICommand
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewCommand"/> class.
        /// </summary>
        /// <param name="executeAction"><see cref="Action"/> to execute on <see cref="Execute"/>.</param>
        public ViewCommand(Action<object> executeAction)
        {
            ExecuteAction = executeAction;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewCommand"/> class.
        /// </summary>
        /// <param name="executeAction"><see cref="Action"/> to execute on <see cref="Execute"/>.</param>
        /// <param name="canExecuteFunction"><see cref="Func&lt;T, TU&gt;"/> to call on <see cref="CanExecute"/>.</param>
        public ViewCommand(Action<object> executeAction, Func<object, bool> canExecuteFunction)
            :this(executeAction)
        {
            CanExecuteFunction = canExecuteFunction;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets <see cref="Action"/> to execute on <see cref="Execute"/>.
        /// </summary>
        public Action<object> ExecuteAction { get; private set; }

        /// <summary>
        /// Gets <see cref="Func&lt;T, TU&gt;"/> to call on <see cref="CanExecute"/>.
        /// </summary>
        public Func<object, bool> CanExecuteFunction { get; private set; }

        #endregion Public properties

        #region Public methods

        /// <summary>
        /// The method calls <see cref="CanExecuteChanged"/>.
        /// </summary>
        public void CallEventCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }

        #endregion Public methods

        #region Implementation of ICommand
        /// <summary>
        /// Occurs when changes occur that affect whether or not the command should execute.
        /// </summary>
        public event EventHandler CanExecuteChanged;

        /// <summary>
        /// Defines the method that determines whether the command can execute in its current state.
        /// </summary>
        /// <param name="parameter">Data used by the command. If the command does not require data to be passed, this object can be set to null.</param>
        /// <returns>true if this command can be executed; otherwise, false.</returns>
        public bool CanExecute(object parameter)
        {
            if (CanExecuteFunction == null)
            {
                return true;
            }
            return CanExecuteFunction(parameter);
        }

        /// <summary>
        /// Defines the method to be called when the command is invoked.
        /// </summary>
        /// <param name="parameter">Data used by the command. If the command does not require data to be passed, this object can be set to null.</param>
        public void Execute(object parameter)
        {
            ExecuteAction?.Invoke(parameter);
        }

        #endregion Implementation of ICommand
    }
}
