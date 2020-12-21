#region Usings
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Model.Properties;
#endregion

namespace AIT.TFS.SyncService.Model.TestReport
{
    /// <summary>
    /// The class implements the interface <see cref="ITfsTestConfiguration"/> for all configurations.
    /// </summary>
    internal class AllTestConfigurations : ITfsTestConfiguration
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="AllTestConfigurations"/> class.
        /// </summary>
        public AllTestConfigurations()
        {
            Id = -1;
            Name = Resources.TestResult_AllConfigurations;
        }

        #endregion Constructors

        #region Public static properties

        /// <summary>
        /// Gets the id of the special configuration.
        /// </summary>
        public static int AllTestConfigurationsId
        {
            get { return -1; }
        }

        #endregion Public static properties

        #region Imlementation of ITfsTestconfiguration

        /// <summary>
        /// Gets the id of the test configuration.
        /// </summary>
        public int Id { get; private set; }

        /// <summary>
        /// Gets the name of test configuration.
        /// </summary>
        public string Name { get; private set; }

        #endregion Implementation of ITfsTestConfiguration
    }
}
