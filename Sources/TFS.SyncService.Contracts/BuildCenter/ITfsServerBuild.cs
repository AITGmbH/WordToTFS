using System;
using Microsoft.TeamFoundation.Build.WebApi;
using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.BuildCenter
{
    /// <summary>
    /// Interface defines functionality of team foundation server build.
    /// </summary>
    public interface ITfsServerBuild
    {
        /// <summary>
        /// Gets the number of the build.
        /// </summary>
        string BuildNumber { get; }

        /// <summary>
        /// Gets the name of build.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets the URI of this build.
        /// </summary>
        Uri Uri { get; }

        /// <summary>
        /// Gets if the type of the current build is either Xaml or Build (Json)
        /// </summary>
        DefinitionType DefinitionType { get; }

        /// <summary>
        /// Gets the Build Quality for Xaml builds, otherwise null
        /// </summary>
        string Quality { get; }

        /// <summary>
        /// Gets a list of tags for the current build respectively an empty list for Xaml builds
        /// </summary>
        List<string> Tags { get; }

        /// <summary>
        /// Gets the datetime when the build has finished or null if the build has not finies, yet
        /// </summary>
        DateTime? FinishTime { get; }
    }
}
