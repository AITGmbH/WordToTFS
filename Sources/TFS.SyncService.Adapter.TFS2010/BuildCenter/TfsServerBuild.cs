#region Usings
using System;
using AIT.TFS.SyncService.Contracts.BuildCenter;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Build.WebApi;
using System.Collections.Generic;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.BuildCenter
{
    /// <summary>
    /// Implements the <see cref="ITfsServerBuild"/>.
    /// </summary>
    internal class TfsServerBuild : ITfsServerBuild
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsServerBuild"/> class.
        /// </summary>
        /// <param name="buildDetail">Build information.</param>
        public TfsServerBuild(IBuildDetail buildDetail)
        {
           BuildDetail = buildDetail;
        }

        public TfsServerBuild(Build build)
        {
           Build = build;
        }

        #endregion Constructors

        #region Internal properties

        /// <summary>
        /// Gets the build detail.
        /// </summary>
        internal IBuildDetail BuildDetail { get; private set; }

        /// <summary>
        /// Gets the build.
        /// </summary>
        internal Build Build { get; private set; }

        #endregion Internal properties

        #region Imlementation of ITfsServerBuild

        /// <summary>
        /// Gets the number of the build.
        /// </summary>
        public string BuildNumber
        {
            get
            {
                if (BuildDetail != null)
                {
                    return BuildDetail.BuildNumber;
                }
                return Build.BuildNumber;
            }
        }

        /// <summary>
        /// Gets the name of build.
        /// </summary>
        public string Name
        {
            get
            {
                if (BuildDetail != null)
                {
                    return BuildDetail.LabelName;
                }
                return $"{Build.Definition.Name}_{Build.BuildNumber}";
            }
        }

        /// <summary>
        /// Gets the URI of this build.
        /// </summary>
        public Uri Uri
        {
            get
            {
                if (BuildDetail != null)
                {
                    return BuildDetail.Uri;
                }
                return Build.Uri;
            }
        }

        /// <summary>
        /// Gets if the type of the current build is either Xaml or Build (Json)
        /// </summary>
        public DefinitionType DefinitionType
        {
            get
            {
                if (BuildDetail != null)
                {
                    return DefinitionType.Xaml;
                }
                return Build.Definition.Type;
            }
        }

        /// <summary>
        /// Gets the Build Quality for Xaml builds, otherwise null
        /// </summary>
        public string Quality
        {
            get
            {
                if (BuildDetail != null)
                {
                    return BuildDetail.Quality;
                }
                return Build.Quality;
            }
        }

        /// <summary>
        /// Gets a list of tags for the current build respectively an empty list for Xaml builds
        /// </summary>
        public List<string> Tags
        {
            get
            {
                if (BuildDetail != null)
                {
                    return new List<string>();
                }
                return Build.Tags;
            }
        }

        /// <summary>
        /// Gets the datetime when the build has finished or null if the build has not finies, yet
        /// </summary>
        public DateTime? FinishTime
        {
            get
            {
                if (BuildDetail != null)
                {
                    return BuildDetail.FinishTime;
                }
                return Build.FinishTime;
            }
        }

        #endregion Imlementation of ITfsServerBuild
    }
}
