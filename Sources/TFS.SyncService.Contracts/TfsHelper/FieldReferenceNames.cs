namespace AIT.TFS.SyncService.Contracts.TfsHelper
{
    /// <summary>
    /// Represents list of field types defined in TFS system.
    /// </summary>
    public static class FieldReferenceNames
    {
        /// <summary>
        /// The unique ID of the work item. Work item IDs are unique across all team projects in a team project collection..
        /// </summary>
        public const string SystemId = "System.Id";

        /// <summary>
        /// The long description of a work item. A text field that provides a more detailed description of the work item than the title provides.
        /// </summary>
        public const string SystemDescription = "System.Description";
        
        /// <summary>
        /// A number that is assigned to the historical revision of a work item. 
        /// </summary>
        public const string SystemRev = "System.Rev";

        /// <summary>
        /// The current state of the work item. The valid values for state are specific for each work item type. The State field supports only the HELPTEXT and READONLY rule types.
        /// </summary>
        public const string SystemState = "System.State";

        /// <summary>
        /// A one-line summary of the work item that helps users distinguish it from other work items in a list.
        /// </summary>
        public const string SystemTitle = "System.Title";

        /// <summary>
        /// The steps required to reproduce unexpected behavior.
        /// </summary>
        public const string TestSteps = "Microsoft.VSTS.TCM.Steps";

        /// <summary>
        /// The stack rank of the work item in the backlog.
        /// </summary>
        public const string StackRank = "Microsoft.VSTS.Common.StackRank";
    }
}
