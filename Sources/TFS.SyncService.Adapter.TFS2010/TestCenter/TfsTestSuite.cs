#region Usings
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements the <see cref="ITfsTestSuite"/>.
    /// </summary>
    public class TfsTestSuite : ITfsTestSuite
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestSuite"/> class.
        /// </summary>
        /// <param name="testSuite">Original test suite - <see cref="IStaticTestSuite"/>.</param>
        public TfsTestSuite(IStaticTestSuite testSuite, ITfsTestPlan associatedTestPlan)
        {

            AssociatedTestPlan = associatedTestPlan;
            OriginalStaticTestSuite = testSuite;
            Id = OriginalStaticTestSuite.Id;
            Title = OriginalStaticTestSuite.Title;
            InitializeChildSuites();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestSuite"/> class.
        /// </summary>
        /// <param name="testSuite">Original test suite - <see cref="IDynamicTestSuite"/>.</param>
        public TfsTestSuite(IDynamicTestSuite testSuite, ITfsTestPlan associatedTestPlan)
        {
            AssociatedTestPlan = associatedTestPlan;
            OriginalDynamicTestSuite = testSuite;
            Id = OriginalDynamicTestSuite.Id;
            Title = OriginalDynamicTestSuite.Title;
            InitializeChildSuites();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestSuite"/> class.
        /// </summary>
        /// <param name="testSuite">The test suite.</param>
        /// <param name="associatedTestPlan">The associated test plan.</param>
        public TfsTestSuite(IRequirementTestSuite testSuite, ITfsTestPlan associatedTestPlan)
        {
            AssociatedTestPlan = associatedTestPlan;
            OriginalRequirementTestSuite = testSuite;
            Id = OriginalRequirementTestSuite.Id;
            Title = OriginalRequirementTestSuite.Title;
            InitializeChildSuites();
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// The associated testPlan
        /// </summary>
        public ITfsTestPlan AssociatedTestPlan
        {
            get;
            private set;
        }
        /// <summary>
        /// Gets the original <see cref="IStaticTestSuite"/>.
        /// </summary>
        public IStaticTestSuite OriginalStaticTestSuite { get; private set; }

        /// <summary>
        /// Gets the original <see cref="IDynamicTestSuite"/>.
        /// </summary>
        public IDynamicTestSuite OriginalDynamicTestSuite { get; private set; }

        /// <summary>
        /// Gets the original <see cref="IDynamicTestSuite"/>.
        /// </summary>
        public IRequirementTestSuite OriginalRequirementTestSuite { get; private set; }

        #endregion Public properties

        #region Private methods

        /// <summary>
        /// The method initializes child test suites.
        /// </summary>
        public void InitializeChildSuites()
        {
            var list = new List<ITfsTestSuite>();
            ITestSuiteCollection childCollection = null;
            if (OriginalStaticTestSuite != null)
                childCollection = OriginalStaticTestSuite.SubSuites;
            if (childCollection != null)
            {
                foreach (var childTestSuite in childCollection)
                {
                    if (childTestSuite is IStaticTestSuite)
                        list.Add(new TfsTestSuite(childTestSuite as IStaticTestSuite, AssociatedTestPlan));
                    if (childTestSuite is IDynamicTestSuite)
                        list.Add(new TfsTestSuite(childTestSuite as IDynamicTestSuite, AssociatedTestPlan));
                    if (childTestSuite is IRequirementTestSuite)
                        list.Add(new TfsTestSuite(childTestSuite as IRequirementTestSuite, AssociatedTestPlan));
                }
            }
            list.Sort((x, y) => string.Compare(x.Title, y.Title));
            TestSuites = list;
        }

        #endregion Private methods

        #region Imlementation of IWordTestSuite

        /// <summary>
        /// Gets the id of the test suite.
        /// </summary>
        public int Id { get; private set; }

        /// <summary>
        /// Gets name of the test suite.
        /// </summary>
        public string Title { get; private set; }

        /// <summary>
        /// Gets all associated test suites.
        /// </summary>
        public IEnumerable<ITfsTestSuite> TestSuites { get; private set; }

        #endregion Imlementation of IWordTestSuite
    }
}
