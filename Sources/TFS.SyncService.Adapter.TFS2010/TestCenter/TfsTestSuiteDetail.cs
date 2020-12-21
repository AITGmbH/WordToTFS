#region Usings
using System;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements <see cref="TfsTestSuiteDetail"/> - detail test suite information.
    /// </summary>
    public class TfsTestSuiteDetail : TfsPropertyValueProvider, ITfsTestSuiteDetail
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestSuiteDetail"/> class.
        /// </summary>
        /// <param name="testSuite">Associated <see cref="TfsTestSuite"/>.</param>
        public TfsTestSuiteDetail(TfsTestSuite testSuite)
        {
            if (testSuite == null)
                throw new ArgumentNullException("testSuite");
            TestSuiteClass = testSuite;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public TfsTestSuite TestSuiteClass { get; private set; }

        /// <summary>
        /// Gets the original <see cref="IStaticTestSuite"/>.
        /// </summary>
        public IStaticTestSuite OriginalStaticTestSuite
        {
            get { return TestSuiteClass.OriginalStaticTestSuite; }
        }

        /// <summary>
        /// Gets the original <see cref="IDynamicTestSuite"/>.
        /// </summary>
        public IDynamicTestSuite OriginalDynamicTestSuite
        {
            get { return TestSuiteClass.OriginalDynamicTestSuite; }
        }

        public IRequirementTestSuite OriginalRequirementTestSuite
        {
            get
            {
                return TestSuiteClass.OriginalRequirementTestSuite;
            }
        }

        #endregion Public properties

        #region Protected override properties

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get
            {
                if (OriginalStaticTestSuite != null)
                    return OriginalStaticTestSuite;
                if (OriginalDynamicTestSuite != null)
                    return OriginalDynamicTestSuite;
                if (OriginalRequirementTestSuite != null)
                    return OriginalRequirementTestSuite;
                throw new ArgumentNullException("AssociatedObject","AssociatedObject was not succesfully initialized!");
            }
        }

        #endregion Protected override properties

        #region Implementation of ITfsTestSuiteDetail

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public ITfsTestSuite TestSuite
        {
            get { return TestSuiteClass; }
        }

        #endregion Implementation of ITfsTestSuiteDetail
    }
}
