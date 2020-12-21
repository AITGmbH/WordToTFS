using System;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012
{
    using Contracts.Adapter;
    using Microsoft.TeamFoundation.Client;
    using System.Windows.Forms;

    /// <summary>
    /// Class implements <see cref="ITeamFoundationServerMaintenance"/>- base team foundation server functionality.
    /// For example selects one team foundation server project from team foundation server
    /// </summary>
    public class TeamFoundationServerMaintenance2010 : ITeamFoundationServerMaintenance
    {
        #region ITeamFoundationServerMaintenance Members

        /// <summary>
        /// Method selects one project from a server by using the built in TeamProjectPicker
        /// </summary>
        /// <param name="server">Server name - valid if the project selected.</param>
        /// <param name="project">Project name - valid if the project selected.</param>
        /// <returns>True if the project selected.</returns>
        public bool SelectOneTfsProject(out string server, out string project)
        {
            server = string.Empty;
            project = string.Empty;


            using (var teamProjectPicker = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false, new UICredentialsProvider()))
            {
                // Return if canceled or if not exactly one project picked
                // The second condition seems pointless (SER)
                if (DialogResult.OK != teamProjectPicker.ShowDialog() ||
                    teamProjectPicker.SelectedProjects.Length != 1)
                {
                    return false;
                }

                server = teamProjectPicker.SelectedTeamProjectCollection.Name;
                project = teamProjectPicker.SelectedProjects[0].Name;
            }
            return server != null;
        }


        /// <summary>
        /// Validate given connection settings and convert the given URL into the server url that can be handled by WordToTFs
        /// </summary>
        /// <param name="defaultServerUrl"></param>
        /// <param name="defaultProjectName"></param>
        /// <param name="serverUrl"></param>
        /// <param name="projectName"></param>
        /// <returns></returns>
        public bool ValidateConnectionSettings(string defaultServerUrl, string defaultProjectName, out string serverUrl, out string projectName)
        {
            try
            {
                var tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(defaultServerUrl, true, false);
                var wiStore = tfs.GetService<WorkItemStore>();
                var projectCollection = wiStore.Projects;

                if (projectCollection.Contains(defaultProjectName))
                {
                    serverUrl = tfs.Name;
                    projectName = defaultProjectName;
                    return true;
                }

            }
            catch (Exception)
            {
                serverUrl = null;
                projectName = null;
                return false;
            }

            serverUrl = null;
            projectName = null;
            return false;
        }
        
        #endregion
    }
}