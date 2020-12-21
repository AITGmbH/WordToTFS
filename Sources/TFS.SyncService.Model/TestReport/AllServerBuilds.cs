#region Usings
using System;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.BuildCenter;
using AIT.TFS.SyncService.Model.Properties;
using Microsoft.TeamFoundation.Build.WebApi;
#endregion

namespace AIT.TFS.SyncService.Model.TestReport
{
    /// <summary>
    /// The class implements the interface <see cref="ITfsServerBuild"/> for all server builds.
    /// </summary>
    public class AllServerBuilds : ITfsServerBuild
    {
        #region Private static fields

        private static AllServerBuilds _empty = null;

        #endregion Private static fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="AllServerBuilds"/> class.
        /// </summary>
        public AllServerBuilds()
        {
            BuildNumber = Resources.TestResult_AllServerBuilds;
            Name = Resources.TestResult_AllServerBuilds;
            Uri = null;
            DefinitionType = DefinitionType.Build | DefinitionType.Xaml;
            Quality = string.Empty;
            Tags = new List<string>();
            FinishTime = null;
        }

        #endregion Constructors

        #region Public static properties

        public static AllServerBuilds Empty
        {
            get
            {
                if (_empty == null)
                {
                    _empty = new AllServerBuilds { BuildNumber = null, Name = null, Uri = null };
                }
                return _empty;
            }
        }

        /// <summary>
        /// Gets the id of the special server build.
        /// </summary>
        public static string AllServerBuildsId
        {
            get { return Resources.TestResult_AllServerBuilds; }
        }

        #endregion Public static properties

        #region Imlementation of ITfsTestconfiguration

        /// <summary>
        /// Gets the number of the build.
        /// </summary>
        public string BuildNumber { get; private set;}

        /// <summary>
        /// Gets the name of build.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the URI of this build.
        /// </summary>
        public Uri Uri { get; private set; }

        /// <summary>
        /// Gets if the type of the current build is either Xaml or Build (Json)
        /// </summary>
        public DefinitionType DefinitionType { get; private set; }

        /// <summary>
        /// Gets the Build Quality for Xaml builds, otherwise an empty string
        /// </summary>
        public string Quality { get; private set; }

        /// <summary>
        /// Gets a list of tags for the current build respectively an empty list for Xaml builds
        /// </summary>
        public List<string> Tags { get; private set; }

        /// <summary>
        /// Gets the datetime when the build has finished or null if the build has not finies, yet
        /// </summary>
        public DateTime? FinishTime { get; private set; }

        #endregion Implementation of ITfsTestConfiguration
    }
}
