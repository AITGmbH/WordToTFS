using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of test case.
    /// </summary>
    public interface ITfsSharedStepDetail : ITfsPropertyValueProvider
    {
        /// <summary>
        /// Gets the base information.
        /// </summary>
        ITfsSharedStep SharedStep { get; }

        /// <summary>
        /// Gets the id of the test case.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets the title of test case.
        /// </summary>
        string Title { get; }

        /// <summary>
        /// Gets the iteration path of test case.
        /// </summary>
        string IterationPath { get; }

        /// <summary>
        /// Gets the area path of test case.
        /// </summary>
        string AreaPath { get; }

        /// <summary>
        /// Gets the number of work item.
        /// </summary>
        int WorkItemId { get; }

        /// <summary>
        /// Special Property that contains all Test Parameters with all values
        /// 
        /// </summary>
        //TODO CHECK if Neccesarry, currently copied from TFSTestCaseDetail, guess is, it is for expanding furthe testcase details
        ITestCaseParameters TestParametersWithAllValues
        {
            get;
            set;
        }
    }
}
