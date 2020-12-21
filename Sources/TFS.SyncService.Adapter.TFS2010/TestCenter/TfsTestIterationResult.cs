#region Usings
using System;
using System.Globalization;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Linq;
using System.Collections.Generic;
using Microsoft.TeamFoundation.TestManagement.Client;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// Wrapper for TestIterationResult that provides TestActionResult-Wrappers
    /// </summary>
    public class TfsTestIterationResult : TfsPropertyValueProvider
    {
        private readonly ITestIterationResult _testIterationResult;
        private readonly ITestCase _testCase;

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestIterationResult"/> class.
        /// </summary>
        /// <param name="testIterationResult">The wrapped test iteration result.</param>
        /// <param name="testCase">The test case used to query the ordering of the test results.</param>
        public TfsTestIterationResult(ITestIterationResult testIterationResult, ITestCase testCase)
        {
            _testIterationResult = testIterationResult;
            _testCase = testCase;
        }

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get
            {
                return _testIterationResult;
            }
        }

        /// <summary>
        /// This property should hide ITestIterationResult.Actions and instead expose a list
        /// of wrappers for those actions. References to shared steps are replaced with the
        /// actions defined in the shared step.
        /// </summary>
        public IList<TfsTestActionResult> Actions
        {
            get
            {
                var actionResults = new List<TfsTestActionResult>();
                var testActionResults = new List<ITestActionResult>(_testIterationResult.Actions);

                // Iterate over all steps in the order DEFINED IN THE TEST CASE and return their results
                for (int i = 0; i < _testCase.Actions.Count; i++)
                {
                    var actionResult = _testIterationResult.FindActionResult(_testCase.Actions[i]);
                    var sharedStepResultReference = actionResult as ISharedStepResult;
                    var testStepResult = actionResult as ITestStepResult;

                    if (testStepResult != null)
                    {
                        testActionResults.Remove(actionResult);
                        actionResults.Add(new TfsTestActionResult(testStepResult, (i + 1).ToString(CultureInfo.InvariantCulture), _testCase.Actions[i]));
                    }

                    // Resolve shared step reference
                    if (sharedStepResultReference != null)
                    {
                        testActionResults.Remove(actionResult);
                        testActionResults.AddRange(sharedStepResultReference.Actions);
 
                        // iterate over the shared steps in the order DEFINED IN THE SHARED STEP
                        for (int sharedStepActionCounter = 0; sharedStepActionCounter < sharedStepResultReference.Actions.Count; sharedStepActionCounter++)
                        {
                            ISharedStep sharedStep;
                            ITestStep sharedStepAction;
                            try
                            {
                                sharedStep = ((ISharedStepReference)_testCase.Actions[i]).FindSharedStep();
                            }
                            catch (DeniedOrNotExistException de)
                            {
                                // I don't know how this happens, but it does. Interestingly, MTM shows the test result with inlined shared steps correctly.
                                throw new InvalidOperationException($"Cannot access shared step ID={sharedStepResultReference.SharedStepId}", de);
                            }

                            // If the shared step is still null, use the current action as test step.
                            if (sharedStep == null)
                            {
                                var sharedStepActionAsTestStep = _testCase.Actions[i] as ITestStep;
                                var sharedStepActionAsTestStepDescription = string.Empty;
                                if (sharedStepActionAsTestStep != null)
                                {
                                    sharedStepActionAsTestStepDescription = sharedStepActionAsTestStep.Description;
                                }
                                throw new InvalidOperationException($"Cannot find result for shared step {sharedStepResultReference.SharedStepId} action {sharedStepActionAsTestStepDescription}");
                            }
                            else
                            {
                                sharedStepAction = (ITestStep)sharedStep.Actions[sharedStepActionCounter];
                            }	

                            var sharedStepActionResult = sharedStepResultReference.Actions[sharedStepActionCounter];
                            testActionResults.Remove(sharedStepActionResult);
                            var stepNumber = $"{i + 1}.{sharedStepActionCounter + 1}";
                            var newActionResult = new TfsTestActionResult(sharedStepActionResult, stepNumber, sharedStepAction);
                            actionResults.Add(newActionResult);
                        }
                    }
                }

                // Unreported results. This should not happen as we are working with the tested test case revision...
                if(testActionResults.Any())
                {
                    throw new InvalidOperationException("Missing definition for ActionResults. The result belongs to an outdated version of the test case.");
                }

                return actionResults;
            }
        }
    }
}
