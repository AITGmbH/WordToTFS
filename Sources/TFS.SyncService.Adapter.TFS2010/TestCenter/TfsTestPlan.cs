using System;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements the <see cref="ITfsTestPlan"/>.
    /// </summary>
    public class TfsTestPlan : ITfsTestPlan
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestPlan"/> class.
        /// </summary>
        /// <param name="testPlan">Original test plan - <see cref="ITestPlan"/>.</param>
        public TfsTestPlan(ITestPlan testPlan)
        {
            if (testPlan == null)
                throw new ArgumentNullException("testPlan");
            OriginalTestPlan = testPlan;
            Id = OriginalTestPlan.Id;
            Name = OriginalTestPlan.Name;
            RootTestSuite = new TfsTestSuite(OriginalTestPlan.RootSuite,this);
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the original test plan.
        /// </summary>
        public ITestPlan OriginalTestPlan { get; private set; }

        #endregion Public properties

        #region Implementation of IWordTestPlan

        /// <summary>
        /// Gets the id of the test plan.
        /// </summary>
        public int Id { get; private set; }

        /// <summary>
        /// Gets the name of test plan.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets root test suite.
        /// </summary>
        public ITfsTestSuite RootTestSuite { get; private set; }

        #endregion Implementation of IWordTestPlan
    }
}
