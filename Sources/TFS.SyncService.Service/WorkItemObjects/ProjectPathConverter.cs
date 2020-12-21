using System;
using System.Globalization;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;

namespace AIT.TFS.SyncService.Service.WorkItemObjects
{
    /// <summary>
    /// Transforms absolute project paths to relative project paths (relative to project name) and vice versa
    /// </summary>
    public class ProjectPathConverter : IConverter
    {
        private readonly string _projectName;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProjectPathConverter"/> class.
        /// </summary>
        /// <param name="projectName">Name of the project.</param>
        public ProjectPathConverter(string projectName)
        {
            _projectName = projectName;
        }

        /// <summary>
        /// The reference field name for which the converter has to be used
        /// </summary>
        /// <exception cref="System.NotImplementedException"></exception>
        public string FieldName
        {
            get
            {
                throw new NotSupportedException();
            }
        }

        /// <summary>
        /// Executes a conversion between the source and the destination value.
        /// The converter allows the manipulation of the values before they are written into the work items.
        /// </summary>
        /// <param name="value">The source value for the conversion.</param>
        /// <param name="direction">Determines the direction of the conversion.</param>
        /// <returns>
        /// The absolute or relative path name, depending on the direction.
        /// </returns>
        /// <exception cref="System.ArgumentNullException">value</exception>
        /// <exception cref="System.ArgumentException"></exception>
        public string Convert(string value, Direction direction)
        {
            Guard.ThrowOnArgumentNull(value, "value");

            if (direction == Direction.TfsToOther)
            {
                if (value.StartsWith(_projectName, StringComparison.OrdinalIgnoreCase))
                {
                    if (value.Equals(_projectName))
                    {
                        return "\\";
                    }
                    return value.Remove(0, _projectName.Length);
                }
                throw new ArgumentException($"The given path is not relative to {_projectName}");
            }
            if (direction == Direction.OtherToTfs)
            {
                return (_projectName + value).TrimEnd('\\');
            }
            return string.Empty;
        }
    }
}