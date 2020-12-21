#region Usings
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Factory;
using System.Linq;
#endregion

namespace AIT.TFS.SyncService.Model.Helper
{
    /// <summary>
    /// Implementation to check if reporting generation should be continued.
    /// </summary>
    public class TestReportingProgressCancellationService : ITestReportingProgressCancellationService
    {
        #region Private fields
        private readonly bool _useProgressService;
        #endregion

        #region Constructor
        /// <summary>
        /// Creates an instance of <see cref="TestReportingProgressCancellationService"/>
        /// </summary>
        /// <param name="useProgressService">Defines if <see cref="IProgressService"/> should be used. w</param>
        public TestReportingProgressCancellationService(bool useProgressService)
        {
            _useProgressService = useProgressService;
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Method to check if cancellation has been triggered by user or an error already has occured.
        /// Returns <c>true</c> if user has not cancelled and no error has occured yet, otherwise <c>false</c>.
        /// </summary>
        /// <returns></returns>
        public bool CheckIfContinue()
        {
            var infoService = SyncServiceFactory.GetService<IInfoStorageService>();
            var progressService = SyncServiceFactory.GetService<IProgressService>();

            var check = (_useProgressService && progressService.ProgressCanceled) || infoService.UserInformation.Any(x => x.Type == UserInformationType.Error);
            return !check;
        }
        #endregion
    }
}
