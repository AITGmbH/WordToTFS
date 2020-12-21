#region Usings
using System;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// Class implements <see cref="ITfsTestRun"/>.
    /// </summary>
    internal class TfsTestRun : ITfsTestRun
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestRun"/> class.
        /// </summary>
        /// <param name="originalTestRun">Original test run - <see cref="ITestRun"/>.</param>
        public TfsTestRun(ITestRun originalTestRun)
        {
            if (originalTestRun == null)
                throw new ArgumentNullException("originalTestRun");
            OriginalTestRun = originalTestRun;
            Id = OriginalTestRun.Id;
            TestSettingsId = OriginalTestRun.TestSettingsId;
            Title = OriginalTestRun.Title;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the original test run - <see cref="ITestRun"/>.
        /// </summary>
        public ITestRun OriginalTestRun { get; private set; }

        #endregion Public properties

        #region Implementation of ITestRun

        /// <summary>
        /// Gets id of the test run.
        /// </summary>
        public int Id { get; private set; }

        /// <summary>
        /// Gets id of test settings.
        /// </summary>
        public int TestSettingsId { get; private set; }

        /// <summary>
        /// Gets title of the test run.
        /// </summary>
        public string Title { get; private set; }

        #endregion Implementation of ITestRun
    }
}
