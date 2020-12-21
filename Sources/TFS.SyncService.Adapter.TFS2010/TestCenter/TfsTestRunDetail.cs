#region Usings
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements <see cref="ITfsTestRunDetail"/>.
    /// </summary>
    internal class TfsTestRunDetail : TfsPropertyValueProvider, ITfsTestRunDetail
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestRunDetail"/> class.
        /// </summary>
        /// <param name="testRun">Associated <see cref="TfsTestRun"/>.</param>
        public TfsTestRunDetail(TfsTestRun testRun)
        {
            TestRunClass = testRun;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the associated <see cref="TfsTestRun"/>.
        /// </summary>
        public TfsTestRun TestRunClass { get; private set; }

        /// <summary>
        /// Gets the original <see cref="ITestRun"/>.
        /// </summary>
        public ITestRun OriginalTestRun
        {
            get { return TestRunClass.OriginalTestRun; }
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
                return OriginalTestRun;
            }
        }

        #endregion Protected override properties

        #region Implementation of ITfsTestRunDetail

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public ITfsTestRun TestRun
        {
            get { return TestRunClass; }
        }

        #endregion Implementation of ITfsTestRunDetail
    }
}
