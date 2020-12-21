#region Usings
using System;
using System.Collections.Generic;
using System.Globalization;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements <see cref="ITfsTestCaseDetail"/> - detail information about test case.
    /// </summary>
    public class TfsTestCaseDetail : TfsPropertyValueProvider, ITfsTestCaseDetail
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestCaseDetail"/> class.
        /// </summary>
        /// <param name="testCase">Associated <see cref="ITfsTestCase"/>.</param>
        /// <param name="expandSharedSteps"> When true, shared steps will be expanded otherwise only the shared steps itself is used (default behavior).</param>
        public TfsTestCaseDetail(TfsTestCase testCase, bool expandSharedSteps = false)
        {
            if (testCase == null)
                throw new ArgumentNullException("testCase");

            ExpandSharedSteps = expandSharedSteps;

            TestCaseClass = testCase;
            Id = TestCase.Id;
            Title = TestCaseClass.OriginalTestCase.Title;
            if (TestCaseClass.OriginalTestCase.WorkItem != null)
            {
                IterationPath = TestCaseClass.OriginalTestCase.WorkItem.IterationPath;
                AreaPath = TestCaseClass.OriginalTestCase.WorkItem.AreaPath;
                WorkItemId = TestCaseClass.OriginalTestCase.WorkItem.Id;
            }

        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public TfsTestCase TestCaseClass { get; private set; }

        /// <summary>
        /// Gets the original test case - <see cref="ITestCase"/>.
        /// </summary>
        public ITestCase OriginalTestCase
        {
            get { return TestCaseClass.OriginalTestCase; }
        }

        #endregion Public properties

        #region Protected override properties

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get { return OriginalTestCase; }
        }

        #endregion Protected override properties

        #region Implementation of ITfsTestCaseDetail

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public ITfsTestCase TestCase
        {
            get { return TestCaseClass; }
        }

        /// <summary>
        /// Gets the id of the test case.
        /// </summary>
        public int Id { get; private set; }

        /// <summary>
        /// Gets the title of test case.
        /// </summary>
        public string Title { get; private set; }

        /// <summary>
        /// Gets the iteration path of test case.
        /// </summary>
        public string IterationPath { get; private set; }

        /// <summary>
        /// Gets the area path of test case.
        /// </summary>
        public string AreaPath { get; private set; }

        /// <summary>
        /// Gets the number of work item.
        /// </summary>
        public int WorkItemId { get; private set; }

        /// <summary>
        /// Gets a value indicating whether shared steps will be expanded.
        /// </summary>
        /// <value>
        ///   <c>true</c> if [expand shared steps]; otherwise, <c>false</c>.
        /// </value>
        public bool ExpandSharedSteps { get; private set; }

        /// <summary>
        /// Special Property that contains all Test Parameters with all values
        /// 
        /// </summary>
        public ITestCaseParameters TestParametersWithAllValues
        {
            get;
            set;
        }

        #endregion Implementation of ITfsTestCaseDetail

        /// <summary>
        /// This property should hide ITestIteration.Actions and instead expose a list of wrappers for those actions
        /// </summary>
        public IList<TfsTestAction> Actions
        {
            get
            {
                var stepCounter = 1;
                var list = new List<TfsTestAction>();

                //TODO USE CONFIG for following:
                //TODO MIS 6.7.2015: Clean Up following logig x.y and SharedSteps Bold y/n

                foreach (var action in TestCaseClass.OriginalTestCase.Actions)
                {
                    var sharedStepReference = action as ISharedStepReference;

                    if (sharedStepReference != null)
                    {
                        //Use sub steps for shared steps, e.g. 3.0 main step and 3.1 sub step
                        if (ExpandSharedSteps)
                        {
                            var innerStepCounter = 1;
                            if (innerStepCounter == 1)
                            {
                                var testHeading = sharedStepReference.FindSharedStep().CreateTestStep();
                                testHeading.Title = sharedStepReference.FindSharedStep().Title;

                                list.Add(new TfsTestAction(testHeading, $"{stepCounter}.0", false));
                            }
                            foreach (var sharedAction in sharedStepReference.FindSharedStep().Actions)
                            {
                                list.Add(new TfsTestAction(sharedAction, $"{stepCounter}.{innerStepCounter++}", false));
                            }

                            if (innerStepCounter > 1)
                            {
                                // At least one inner step has been inserted.
                                stepCounter++;
                            }
                        }
                        else
                        {
                            var testHeading = sharedStepReference.FindSharedStep().CreateTestStep();
                            testHeading.Title = sharedStepReference.FindSharedStep().Title;
                            list.Add(new TfsTestAction(testHeading, (stepCounter++).ToString(CultureInfo.InvariantCulture), true));
                        }
                    }
                    else if (action is ITestStep)
                    {
                        list.Add(new TfsTestAction(action, (stepCounter++).ToString(CultureInfo.InvariantCulture), false));
                    }
                }
                return list;
            }
        }

    }
}
