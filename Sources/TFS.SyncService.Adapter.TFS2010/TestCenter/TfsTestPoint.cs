using System;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    public class TfsTestPoint : ITfsTestPoint
    {

        public TfsTestPoint(ITestPoint originalTestPoint)
        {
            if (originalTestPoint == null) throw new ArgumentNullException("originalTestCase");
            OriginalTestPoint = originalTestPoint;
            Id = originalTestPoint.Id;
        }

        /// <summary>
        /// Gets the original test case - <see cref="ITestCase"/>.
        /// </summary>
        public ITestPoint OriginalTestPoint { get; private set; }

        /// <summary>
        /// Gets the id of test point
        /// </summary>
        public int Id
        {
            get;
            private set;
        }
    }

    
}
