namespace AIT.TFS.SyncService.Contracts.TemplateManager
{
    /// <summary>
    /// The interface defines functionality of one template in template bundle.
    /// </summary>
    public interface ITemplate
    {
        /// <summary>
        /// Gets the show name of template in template bundle.
        /// </summary>
        string ShowName
        {
            get;
        }

        /// <summary>
        /// Gets or the name of the project this template is restricted to. This template should only be
        /// available if the user is connected to a project with the given name. If set to null or empty string
        /// the template is always available.
        /// </summary>
        string ProjectName
        {
            get;
        }

        /// <summary>
        /// Gets or the name of the project collection this template is restricted to. This template should only be
        /// available if the user is connected to a project with the given name. If set to null or empty string
        /// the template is always available.
        /// </summary>
        string ProjectCollectionName
        {
            get;
        }


        /// <summary>
        /// Gets the name of the server this template is restricted to. This template should only be
        /// available if the user is connected to a server with the given name. If set to null or empty string
        /// the template is always available.
        /// </summary>
        string ServerName
        {
            get;
        }

        /// <summary>
        /// Gets or sets the state of template - describes state of template.
        /// </summary>
        TemplateState TemplateState
        {
            get;
        }
    }
}
