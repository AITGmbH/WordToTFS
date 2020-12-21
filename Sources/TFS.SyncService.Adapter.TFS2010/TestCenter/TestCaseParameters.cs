using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.TestCenter;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{

    /// <summary>
    /// Class that represents the parameters of a test case
    /// </summary>
    public class TestCaseParameters : ITestCaseParameters
    {

        /// <summary>
        /// Shows all paramters for a testcase
        /// </summary>
        public TestCaseParameters()
        {
            Iterations = new List<ITestCaseIterations>();
          
        }

        /// <summary>
        /// Gets all parameters
        /// </summary>
        public string AllParameters
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the interations
        /// </summary>
        public List<ITestCaseIterations> Iterations
        {
            get;
            private set;
        }
    }
}
