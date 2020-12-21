using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Service.Configuration.Serialization;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// The class implements the interface <see cref="IConfigurationFieldAssignment"/>.
    /// The class contains information about field assignment for linked work item.
    /// </summary>
    public class ConfigurationFieldAssignment : IConfigurationFieldAssignment
    {
        #region Constructors
        /// <summary>
        /// Constructor creates an instance of the class <see cref="ConfigurationFieldAssignment"/>.
        /// </summary>
        /// <param name="copyFrom">The parameter defines the values of properties in the instance of the class.</param>
        public ConfigurationFieldAssignment(MappingFieldAssignment copyFrom)
        {
            if (copyFrom == null)
                return;
            ReferenceName = copyFrom.Name;
            SourceMappingName = copyFrom.MappingName;
            FieldPosition = copyFrom.FieldPosition;
            FieldValueType = copyFrom.FieldValueType;
            Direction = copyFrom.Direction;
        }

        #endregion Constructors

        #region Implementation of IConfigurationFieldAssignment
        /// <summary>
        /// The field name of the destination field
        /// </summary>
        public string ReferenceName { get; private set; }

        /// <summary>
        /// The field name of the source field
        /// </summary>
        public string SourceMappingName { get; private set; }

        /// <summary>
        /// Gets the position of the field in the definition of linked work item.
        /// </summary>
        public FieldPositionType FieldPosition { get; private set; }

        /// <summary>
        /// Determines whether the exported value should be plain text or html
        /// </summary>
        public FieldValueType FieldValueType { get; private set; }

        /// <summary>
        /// Determines the copy direction of the field
        /// </summary>
        public Direction Direction { get; private set; }
        #endregion Implementation of IConfigurationFieldAssignment
    }
}
