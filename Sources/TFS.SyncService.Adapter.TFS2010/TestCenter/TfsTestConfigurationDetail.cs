using System;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// The class implements <see cref="ITfsTestConfigurationDetail"/> - detail information about test configuration.
    /// </summary>
    public class TfsTestConfigurationDetail : TfsPropertyValueProvider, ITfsTestConfigurationDetail
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestConfigurationDetail"/> class.
        /// </summary>
        /// <param name="testConfiguration">Associated <see cref="TfsTestConfiguration"/>.</param>
        public TfsTestConfigurationDetail(TfsTestConfiguration testConfiguration)
        {
            if (testConfiguration == null)
                throw new ArgumentNullException("testConfiguration");
            TestConfigurationClass = testConfiguration;
        }

        #endregion Constructors

        #region Public properties

        /// <summary>
        /// Gets the base information.
        /// </summary>
        public TfsTestConfiguration TestConfigurationClass { get; private set; }

        /// <summary>
        /// Gets the original test configuration - <see cref="ITestConfiguration"/>.
        /// </summary>
        public ITestConfiguration OriginalTestConfiguration
        {
            get { return TestConfigurationClass.OriginalTestConfiguration; }
        }

        #endregion Public properties

        #region Protected override properties

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get { return OriginalTestConfiguration; }
        }

        #endregion Protected override properties

        #region Implementation of ITfsTestConfigurationDetail
        
        /// <summary>
        /// Gets the base information.
        /// </summary>
        public ITfsTestConfiguration TestConfiguration
        {
            get { return TestConfigurationClass; }
        }

        #endregion Implementation of ITfsTestConfigurationDetail
    }
}
