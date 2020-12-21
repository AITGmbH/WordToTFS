namespace AIT.TFS.SyncService.Contracts
{
    /// <summary>
    /// Static class used as container of various constants.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Constant defines the name of company.
        /// </summary>
        public const string ApplicationCompany = "AIT";

        /// <summary>
        /// Constant defines the name of application.
        /// </summary>
        public const string ApplicationName = "WordToTFS2010";

        /// <summary>
        /// Subfolder with all help related files.
        /// </summary>
        public const string HelpFileSubfolder = "Help";

        /// <summary>
        /// Subfolder with all mapping bundle files.
        /// </summary>
        public const string MappingBundleSubfolder = "TemplateCache";

        /// <summary>
        /// Subfolder with the default AIT template bundle.
        /// </summary>
        public const string DefaultTemplateBundleSubfolder = "Templates";

        /// <summary>
        /// Name of the log file
        /// </summary>
        public const string LogFileName = "WordToTFS2010.Log.txt";

        /// <summary>
        /// Name of the HierarchyLevel property used to assign different templates to work items
        /// </summary>
        public const string NameOfHierarchyLevelProperty = "HierarchyLevel";
    }
}