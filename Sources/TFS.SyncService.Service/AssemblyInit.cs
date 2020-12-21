using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Service.Configuration;
using AIT.TFS.SyncService.Service.InfoStorage;
using AIT.TFS.SyncService.Service.Progress;

namespace AIT.TFS.SyncService.Service
{
    /// <summary>
    /// Class implements the behavior of initialization of one assembly for 'ClickOnce Deployment'.
    /// </summary>
    public sealed class AssemblyInit : IAssemblyInit
    {
        /// <summary>
        /// Holds the singleton instance.
        /// </summary>
        private static AssemblyInit _instance;

        /// <summary>
        /// Gets the singleton instance.
        /// </summary>
        public static IAssemblyInit Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new AssemblyInit();
                return _instance;
            }
        }

        #region IAssemblyInit Members

        /// <summary>
        /// Method initializes the assembly.
        /// </summary>
        public void Init()
        {
            SyncServiceFactory.RegisterService<IInfoStorageService>(new InfoStorageService());
            var confService = new ConfigurationService();
            SyncServiceFactory.RegisterService<IConfigurationService>(confService);
            SyncServiceFactory.RegisterService<IProgressService>(new ProgressService());
            var systemVariableService = new SystemVariableService();
            SyncServiceFactory.RegisterService<ISystemVariableService>(systemVariableService);
            SyncServiceFactory.RegisterService<IWorkItemSyncService>(new WorkItemSyncService(systemVariableService));
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            // delete temporary directory
            // TODO what was deleted here?
            //var tempPath = TemporaryPathHelper.GetTempPath();
            //if (Directory.Exists(tempPath)) Directory.Delete(tempPath, true);
        }

        #endregion
    }
}