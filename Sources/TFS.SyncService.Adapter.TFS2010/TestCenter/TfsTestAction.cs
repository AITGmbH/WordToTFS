namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    using Microsoft.TeamFoundation.TestManagement.Client;

    /// <summary>
    /// Wrapper that provides a StepNumber-Property the templates can be bound against. The step number
    /// is the index within the TestActionResultCollection of a TestIterationResult
    /// </summary>
    public class TfsTestAction : TfsPropertyValueProvider
    {
        private readonly ITestAction _testAction;

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestAction"/> class.
        /// </summary>
        /// <param name="testAction">The wrapped test action.</param>
        /// <param name="stepNumber"
        /// >The step number as defined in the test case.</param>
        /// <param name="isSharedStepTitle">Determines that this action is the title of the shared step.</param>
        public TfsTestAction(ITestAction testAction, string stepNumber, bool isSharedStepTitle)
        {
            _testAction = testAction;
            StepNumber = stepNumber;
            IsSharedStepTitle = isSharedStepTitle;

        }


        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get
            {
                return _testAction;
            }
        }

        /// <summary>
        /// Gets the StepNumber of this test action
        /// </summary>
        public string StepNumber { get; private set; }

        /// <summary>
        /// Gets a value indicating whether this instance is shared step title.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is shared step title; otherwise, <c>false</c>.
        /// </value>
        public bool IsSharedStepTitle { get; private set; }

    }
}
