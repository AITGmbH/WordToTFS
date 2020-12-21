namespace AIT.TFS.SyncService.Service.Configuration.Model
{
    /// <summary>
    /// Class holds the information about one mapping bundle that can be used in plug in.
    /// </summary>
    internal class MappingInfo
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MappingInfo"/> class.
        /// </summary>
        /// <param name="showName">Name of the mapping bundle.</param>
        /// <param name="mappingFileName">Mapping file.</param>
        public MappingInfo(string showName, string mappingFileName)
        {
            ShowName = showName;
            MappingFileName = mappingFileName;
        }

        /// <summary>
        /// Gets or sets the name of the mapping bundle.
        /// </summary>
        public string ShowName { get; set; }

        /// <summary>
        /// Gets or sets the mapping file.
        /// </summary>
        public string MappingFileName { get; set; }
    }
}