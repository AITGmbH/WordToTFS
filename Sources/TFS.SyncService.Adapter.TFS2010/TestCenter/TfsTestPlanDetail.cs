using System;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements <see cref="ITfsTestPlanDetail"/> - detail information about test plan.
    /// </summary>
    public class TfsTestPlanDetail : TfsPropertyValueProvider, ITfsTestPlanDetail
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestPlanDetail"/> class.
        /// </summary>
        /// <param name="testPlan">Associated <see cref="TfsTestPlan"/>.</param>
        public TfsTestPlanDetail(TfsTestPlan testPlan)
        {
            if (testPlan == null)
                throw new ArgumentNullException("testPlan");
            TestPlanClass = testPlan;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the base information
        /// </summary>
        public TfsTestPlan TestPlanClass { get; private set; }

        /// <summary>
        /// Gets the original test plan.
        /// </summary>
        public ITestPlan OriginalTestPlan
        {
            get { return TestPlanClass.OriginalTestPlan; }
        }

        #endregion Public properties

        #region Protected override properties

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get { return OriginalTestPlan; }
        }

        #endregion Protected override properties

        #region Implementation of ITfsTestPlanDetail

        /// <summary>
        /// Gets the base information
        /// </summary>
        public ITfsTestPlan TestPlan
        {
            get { return TestPlanClass; }
        }

        #endregion Implementation of ITfsTestPlanDetail
    }
}
