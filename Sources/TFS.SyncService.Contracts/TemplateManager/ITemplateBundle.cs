namespace AIT.TFS.SyncService.Contracts.TemplateManager
{
    /// <summary>
    /// Interface defines the functionality of 
    /// </summary>
    public interface ITemplateBundle
    {
        /// <summary>
        /// Gets the show name of template bundle.
        /// </summary>
        string ShowName
        {
            get;
        }

        /// <summary>
        /// Gets or sets the type of template bundle - type defines the origin of the template bundle.
        /// </summary>
        TemplateBundleType TemplateBundleType
        {
            get;
        }

        /// <summary>
        /// Gets or sets the state of whole template bundle - describes state of whole template bundle.
        /// </summary>
        TemplateBundleStates TemplateBundleState
        {
            get;
        }
    }
}
