using System;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements <see cref="IConfigurationTestUriLink"/>.
    /// </summary>
    public class ConfigurationTestUriLink : IConfigurationTestUriLink
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationTestUriLink"/> class.
        /// </summary>
        /// <param name="uriLink">Associated configuration.</param>
        public ConfigurationTestUriLink(UriLink uriLink)
        {
            if (uriLink == null)
                throw new ArgumentNullException("uriLink");
            Uri = uriLink.Uri;
        }

        #endregion Constructors

        #region Implementation of IConfigurationTestUriLink

        /// <summary>
        /// Gets or sets the property to evaluate and gets the common uri.
        /// </summary>
        public string Uri { get; private set; }

        #endregion Implementation of IConfigurationTestUriLink
    }
}
