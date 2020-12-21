#region Usings
using System;
using AIT.TFS.SyncService.Contracts.Configuration;
#endregion

namespace TFS.Test.Common
{
    /// <summary>
    /// Server Configuration
    /// </summary>
    public class ServerConfiguration
    {
        /// <summary>
        /// Gets or sets the server name of the project used for testing
        /// </summary>
        public string TeamProjectCollectionUrl
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the project name of the project used for testing
        /// </summary>
        public string TeamProjectName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the reference name of the field used for HTML tests
        /// </summary>
        public string HtmlFieldReferenceName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the id of the work item used for HTML integration tests
        /// </summary>
        public int HtmlRequirementId
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the id of the work item used for SynchronizationState testing.
        /// This work item must have its System.AreaPath changed in the last change and nothing else.
        /// </summary>
        public int SynchronizedWorkItemId
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the configuration to be used with this server
        /// </summary>
        public IConfiguration Configuration
        {
            get;
            set;
        }


        /// <summary>
        /// the computer name
        /// </summary>
        public string ComputerName
        {
            get;
            set;
        }

        /// <summary>
        /// the computer name
        /// </summary>
        public string FullyQualifiedComputerName
        {
            get;
            set;
        }


        /// <summary>
        /// The collection name
        /// </summary>
        public string TeamProjectCollectionName
        {
            get;
            set;
        }
    }
}
