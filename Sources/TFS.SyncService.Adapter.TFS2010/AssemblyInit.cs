namespace AIT.TFS.SyncService.Adapter.TFS2012
{
    using Contracts;
    using Contracts.Adapter;
    using Factory;

    /// <summary>
    /// Class implements the behavior of initialization of one assembly for 'ClickOnce Deployment'.
    /// </summary>
    public sealed class AssemblyInit : IAssemblyInit
    {
        /// <summary>
        /// Holds the singleton instance
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
            SyncServiceFactory.RegisterService<ITeamFoundationServerMaintenance>(
                new TeamFoundationServerMaintenance2010());
            SyncServiceFactory.RegisterService<ITfsAdapterCreator>(new AdapterCreator());
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            // Do nothing at this moment
        }

        #endregion
    }
}