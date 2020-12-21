using AIT.TFS.SyncService.Adapter.Word2007.TestCenter;
using AIT.TFS.SyncService.Contracts.Adapter;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace AIT.TFS.SyncService.Adapter.Word2007
{
    /// <summary>
    /// Class implements the behavior for creator of an adapter for word 2007.
    /// </summary>
    internal class AdapterCreator : IWord2007AdapterCreator
    {
        #region Implementation of IWord2007AdapterCreator

        /// <summary>
        /// Creates an instance of word 2007 adapter.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">Configuration to use</param>
        /// <returns>Created adapter.</returns>
        public IWordSyncAdapter CreateTableAdapter(object document, IConfiguration configuration)
        {
            return new Word2007TableSyncAdapter(document as Document, configuration);
        }

        /// <summary>
        /// Creates an instance of word 2007 test report adapter.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">Configuration to use.</param>
        /// <returns>Created adapter.</returns>
        public IWord2007TestReportAdapter CreateTestReportAdapter(object document, IConfiguration configuration)
        {
            return new Word2007TestReportAdapter(document as Document, configuration);
        }

        /// <summary>
        /// Creates an instance of word 2007 test report adapter.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">Configuration to use.</param>
        /// <param name="testSuite">Test suite.</param>
        ///  <param name="allTestCases">All test cases that test suite contains.</param>
        /// <returns>Created adapter.</returns>
        public IWord2007TestReportAdapter CreateTestReportAdapter(object document, IConfiguration configuration, ITfsTestSuite testSuite, IList<ITfsTestCaseDetail> allTestCases)
        {
            return new Word2007TestReportAdapter(document as Document, configuration, testSuite, allTestCases);
        }

        #endregion Implementation of IWord2007AdapterCreator
    }
}