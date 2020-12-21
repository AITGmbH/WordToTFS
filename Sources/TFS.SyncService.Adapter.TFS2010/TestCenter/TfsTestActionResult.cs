using AIT.TFS.SyncService.Contracts.Exceptions;
using System.Linq;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    using Microsoft.TeamFoundation.TestManagement.Client;
    using Factory;
    using Properties;

    /// <summary>
    /// Wrapper that provides a StepNumber-Property the templates can be bound against. The step number
    /// is the index within the TestActionResultCollection of a TestIterationResult
    /// </summary>
    public class TfsTestActionResult : TfsPropertyValueProvider
    {
        private readonly ITestActionResult _testActionResult;
        private readonly ITestAction _testAction;

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestActionResult"/> class.
        /// </summary>
        /// <param name="testActionResult">The wrapped test action result.</param>
        /// <param name="stepNumber">The step number of this result as defined in the tested test case.</param>
        /// <param name="testAction">The original test action.</param>
        public TfsTestActionResult(ITestActionResult testActionResult, string stepNumber, ITestAction testAction)
        {
            _testAction = testAction;
            _testActionResult = testActionResult;
            StepNumber = stepNumber;
        }

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get
            {
                return _testActionResult;
            }
        }

        /// <summary>
        /// Gets the StepNumber of this test action
        /// </summary>
        public string StepNumber { get; private set; }

        /// <summary>
        /// Gets the expected result of the related test step where parameters are
        /// replaced by the values of the tested iteration.
        /// </summary>
        public string ExpectedResult
        {
            get
            {
                var testStep = _testAction as ITestStep;
                var testStepResult = _testActionResult as ITestStepResult;

                if(testStep == null || testStepResult == null)
                {
                    SyncServiceTrace.W(Resources.ExpectedResultIsOnlyImplementedForTestStepResults);
                    return string.Empty;
                }

                return ReplaceParameters(testStep.ExpectedResult, testStepResult.Parameters);
            }
        }

        /// <summary>
        /// Gets the title of the related test step where parameters are
        /// replaced by the values of the tested iteration.
        /// </summary>
        public string Title
        {
            get
            {
                var testStep = _testAction as ITestStep;
                var testStepResult = _testActionResult as ITestStepResult;

                if (testStep == null || testStepResult == null)
                {
                    SyncServiceTrace.W(Resources.ExpectedResultIsOnlyImplementedForTestStepResults);
                    return string.Empty;
                }

                return ReplaceParameters(testStep.Title, testStepResult.Parameters);
            }
        }

        /// <summary>
        /// Gets the related TestAction
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Justification = "User needs immediate feedback when his configuration is broken.")]
        public TfsTestAction TestStep
        {
            get
            {
                if (!(_testAction is ITestStep))
                {
                    throw new ConfigurationException("The extension property 'TestStep' is only implemented for TestStepResults");
                }

                return new TfsTestAction(_testAction, StepNumber, false);
            }
        }


        /// <summary>
        /// Replaces all parameters in a parameterized text.
        /// </summary>
        /// <param name="parameterizedText">The parameterized text.</param>
        /// <param name="parameters">The parameters.</param>
        /// <returns>Text where all parameter references have been replaced with expected values</returns>
        private static string ReplaceParameters(ParameterizedString parameterizedText, TestResultParameterCollection parameters)
        {
            if (parameters.Count > 0)
            {
                SyncServiceTrace.D(Resources.ReplacingParameters, parameterizedText, string.Join(",", parameters.Select(x => x.Name + "=" + x.ExpectedValue)));
                foreach (var parameter in parameters)
                {
                    var oldParameterizedText = parameterizedText.ToPlainText();
                    parameterizedText = parameterizedText.ReplaceParameter(parameter.Name, parameter.ExpectedValue);

                    // I don't know why this fails sometimes, maybe because of special characters in the string like "foo: @paramter"
                    if (parameterizedText.ToPlainText().Equals(oldParameterizedText) && parameterizedText.ParameterNames.Contains(parameter.Name))
                    {
                        // manually replace parameter
                        SyncServiceTrace.W(Resources.ParameterizedStringFailed, parameter.Name, parameterizedText.ToPlainText(), parameter.ExpectedValue);
                        parameterizedText = new ParameterizedString(parameterizedText.ToPlainText().Replace("@" + parameter.Name, parameter.ExpectedValue));
                    }
                }
            }
            return parameterizedText.ToPlainText();
        }


    }
}
