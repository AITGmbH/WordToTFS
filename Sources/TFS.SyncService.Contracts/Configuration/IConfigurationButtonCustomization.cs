
namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Interface defines functionality for one field in work item.
    /// </summary>
    public interface IConfigurationButtonCustomization
    {
        /// <summary>
        /// The name of the button which should be renamed
        /// </summary>
        string Name
        {
            get;
            set;
        }

        /// <summary>
        /// The text of which will be displayed
        /// </summary>
        string Text
        {
            get;
            set;
        }
    }
}