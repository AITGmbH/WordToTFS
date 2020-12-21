
using System.Collections.Generic;


namespace AIT.TFS.SyncService.Contracts.TestCenter
{

    /// <summary>
    /// A property that allows the display of parameters for testcases
    /// </summary>
    public interface ITestCaseParameters
    {

        /// <summary>
        /// A string representation of all parameters of the testcase
        /// </summary>
        string AllParameters
        {
            get;
            set;
        }


        /// <summary>
        /// A list of all iterations for the testcase
        /// </summary>
        List<ITestCaseIterations> Iterations
        {
            get;
        } 


    }

}
