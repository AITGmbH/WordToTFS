using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Contracts.WorkItemObjects
{
    /// <summary>
    /// Interface defines functionality of a converter to convert field from one system to another.
    /// </summary>
    public interface IConverter
    {
        /// <summary>
        /// The reference field name for which the converter has to be used
        /// </summary>
        string FieldName { get; }

        /// <summary>
        /// Executes a conversion between the source and the destination value. 
        /// The converter allows the manipulation of the values before they are written into the work items.
        /// </summary>
        /// <param name="value">The source value for the conversion.</param>
        /// <param name="direction">Determines the direction of the conversion.</param>
        /// <returns>Always a valid string with the converted value.</returns>
        string Convert(string value, Direction direction);
    }
}