using System.Diagnostics.CodeAnalysis;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace AIT.TFS.SyncService.Contracts.Adapter
{
    /// <summary>
    /// Interface defines functionality for maintenance of team foundation server.
    /// </summary>
    public interface ITeamFoundationServerMaintenance
    {
        /// <summary>
        /// Method selects one project from one server by using a build in dialogue.
        /// </summary>
        /// <param name="server">Server URL - valid if the project selected.</param>
        /// <param name="project">Project name - valid if the project selected.</param>
        /// <returns>True if the project selected.</returns>
        [SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters")]
        bool SelectOneTfsProject(out string server, out string project);
        /// <summary>
        /// Validate given connection settings and convert the given URL into the server url that can be handled by WordToTFs
        /// </summary>
        /// <summary>
        /// Validate given connection settings and convert the given URL into the server url that can be handled by WordToTFs
        /// </summary>
        /// <param name="defaultServerUrl"></param>
        /// <param name="defaultProjectName"></param>
        /// <param name="serverUrl"></param>
        /// <param name="projectName"></param>
        /// <returns></returns>
        bool ValidateConnectionSettings(string defaultServerUrl, string defaultProjectName, out string serverUrl, out string projectName);
    }
}