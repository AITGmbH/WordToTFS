using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.TestCenter;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{

    /// <summary>
    /// A TestCase Iteration that holds the parameters
    /// </summary>
    public class TestCaseIterations : ITestCaseIterations
    {
        public int IterationNumber
        {
            get;
            set;
        }

        /// <summary>
        /// 
        /// </summary>
        public IEnumerable<IParameter> Parameters
        {
            get;
            set;
        }

    }

    /// <summary>
    /// The Parameter that holds the value
    /// </summary>
    public class Parameter : IParameter
    {
        public string ParameterValue
        {
            get;
            set;
        }
    }
}
