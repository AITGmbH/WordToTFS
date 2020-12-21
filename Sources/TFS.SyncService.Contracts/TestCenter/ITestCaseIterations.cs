
using System.Collections.Generic;


namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    public interface ITestCaseIterations
    {

        /// <summary>
        /// The interface for the test case Iteration
        /// </summary>
        int IterationNumber
        {
            get;
        }

        /// <summary>
        /// The interface for the paramters of a test case iteration
        /// </summary>
        IEnumerable<IParameter> Parameters
        {
            get;
        }

    }

    /// <summary>
    /// The Interface that defines a Parameter
    /// </summary>
    public interface IParameter
    {
        string ParameterValue
        {
            get;
        }

        
    }
}
