using AIT.TFS.SyncService.Contracts.TestCenter;

namespace AIT.TFS.SyncService.Contracts.Adapter
{
    using Configuration;
    using System.Collections.Generic;

    /// <summary>
    /// Interface defines the behavior for creator of an adapter for word 2007.
    /// </summary>
    public interface IWord2007AdapterCreator
    {
        /// <summary>
        /// Creates an instance of word 2007 adapter.
        /// </summary>
        /// <param name="document">word document to work with.</param>
        /// <param name="configuration">Configuration to use</param>
        /// <returns>Created adapter.</returns>
        IWordSyncAdapter CreateTableAdapter(object document, IConfiguration configuration);

        /// <summary>
        /// Creates an instance of word 2007 test report adapter.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">Configuration to use.</param>
        /// <returns>Created adapter.</returns>
        IWord2007TestReportAdapter CreateTestReportAdapter(object document, IConfiguration configuration);

        /// <summary>
        /// Creates an instance of word 2007 test report adapter.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">Configuration to use.</param>
        /// <param name="testSuite">Test suite.</param>
        /// <param name="allTestCases">All test cases that test suite contains.</param>
        /// <returns>Created adapter.</returns>
        IWord2007TestReportAdapter CreateTestReportAdapter(object document, IConfiguration configuration, ITfsTestSuite testSuite, IList<ITfsTestCaseDetail> allTestCases);
    }
}