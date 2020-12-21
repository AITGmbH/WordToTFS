using System;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements the <see cref="ITfsTestConfiguration"/>.
    /// </summary>
    public class TfsTestConfiguration : ITfsTestConfiguration
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestConfiguration"/> class.
        /// </summary>
        /// <param name="originalTestConfiguration">Original test configuration - <see cref="ITestConfiguration"/>.</param>
        public TfsTestConfiguration(ITestConfiguration originalTestConfiguration)
        {
            if (originalTestConfiguration == null)
                throw new ArgumentNullException("originalTestConfiguration");
            OriginalTestConfiguration = originalTestConfiguration;
            Id = OriginalTestConfiguration.Id;
            Name = OriginalTestConfiguration.Name;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the original test configuration - <see cref="ITestConfiguration"/>.
        /// </summary>
        public ITestConfiguration OriginalTestConfiguration { get; private set; }

        #endregion Public properties

        #region Implementation of ITfsTestConfiguration
        /// <summary>
        /// Gets the id of the test configuration.
        /// </summary>
        public int Id { get; internal set; }

        /// <summary>
        /// Gets the name of test configuration.
        /// </summary>
        public string Name { get; internal set; }
        #endregion Implementation of ITfsTestConfiguration
    }
}
