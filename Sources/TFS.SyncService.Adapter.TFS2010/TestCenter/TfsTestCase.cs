#region Usings
using System;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements <see cref="ITfsTestCase"/>.
    /// </summary>
    public class TfsTestCase : ITfsTestCase
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestCase"/> class.
        /// </summary>
        /// <param name="originalTestCase">Original test case - <see cref="ITestCase"/>.</param>
        public TfsTestCase(ITestCase originalTestCase)
        {
            if (originalTestCase == null)
                throw new ArgumentNullException("originalTestCase");
            OriginalTestCase = originalTestCase;
            Id = OriginalTestCase.Id;
            Title = OriginalTestCase.WorkItem.Title;
            TestParametersWithAllValues = new List<TestCaseParameters>();
        }


        public TfsTestCase(ITestCase originalTestCase, ITfsTestSuite testSuiteDetail) : this (originalTestCase)
        {

            AssociatedTestSuiteDetail = testSuiteDetail;
        }
        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the original test case - <see cref="ITestCase"/>.
        /// </summary>
        public ITestCase OriginalTestCase { get; private set; }

        public ITfsTestSuite AssociatedTestSuiteDetail { get; private set; }
        #endregion Public properties

        #region Implementation of ITfsTestCase

        /// <summary>
        /// Gets the id of the test case.
        /// </summary>
        public int Id { get; private set; }

        /// <summary>
        /// Gets the title of the test case.
        /// </summary>
        public string Title { get; private set; }

        public List<TestCaseParameters> TestParametersWithAllValues
        {
            get;
            set;
        }
        #endregion Implementation of ITfsTestCase
    }
}
