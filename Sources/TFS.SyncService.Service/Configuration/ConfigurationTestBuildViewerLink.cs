using System;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationTestBuildViewerLink"/>.
    /// </summary>
    public class ConfigurationTestBuildViewerLink : IConfigurationTestBuildViewerLink
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestBuildViewerLink"/> class.
        /// </summary>
        /// <param name="buildViewerLink">Associated configuration.</param>
        public ConfigurationTestBuildViewerLink(BuildViewerLink buildViewerLink)
        {
            if (buildViewerLink == null)
                throw new ArgumentNullException("buildViewerLink");
            BuildNumber = buildViewerLink.BuildNumber;
        }

        #endregion Constructors

        #region Implementation of IConfigurationTestBuildViewerLink

        /// <summary>
        /// Gets or sets the property to evaluate and gets the build number of the build.
        /// </summary>
        public string BuildNumber { get; private set; }

        #endregion Implementation of IConfigurationTestBuildViewerLink
    }
}
