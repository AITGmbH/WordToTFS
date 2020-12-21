using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// The interface defines the functionality of configuration for field assignment to define linked work item in detail.
    /// </summary>
    public interface IConfigurationFieldAssignment
    {
        /// <summary>
        /// The field name of the destination field
        /// </summary>
        string ReferenceName { get; }

        /// <summary>
        /// The field name of the source field
        /// </summary>
        string SourceMappingName { get; }

        /// <summary>
        /// Gets the position of the field in the definition of linked work item.
        /// </summary>
        FieldPositionType FieldPosition { get; }

        /// <summary>
        /// Determines whether the exported value should be plain text or html
        /// </summary>
        FieldValueType FieldValueType { get; }

        /// <summary>
        /// Determines the copy direction of the field
        /// </summary>
        Direction Direction { get; }
    }
}
