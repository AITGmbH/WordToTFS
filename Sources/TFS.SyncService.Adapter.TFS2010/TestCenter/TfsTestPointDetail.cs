using System;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    class TfsTestPointDetail : TfsPropertyValueProvider, ITfsTestPointDetail
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestCaseDetail"/> class.
        /// </summary>
        /// <param name="testPoint">Associated testpoint</param>
   
        public TfsTestPointDetail(TfsTestPoint testPoint)
        {
            TestPoint = testPoint;
            TestPointClass = testPoint;
            if (testPoint == null) throw new ArgumentNullException("testPoint");
            Id = testPoint.Id;
        }

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public TfsTestPoint TestPointClass { get; private set; }

        /// <summary>
        /// Gets the original test case - <see cref="ITestCase"/>.
        /// </summary>
        public ITestPoint OriginalTestPoint
        {
            get { return TestPointClass.OriginalTestPoint; }
        }

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get { return OriginalTestPoint; }
        }

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public ITfsTestPoint TestPoint
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the id of the test case.
        /// </summary>
        public int Id
        {
            get;
            private set;
        }
    }
}
