namespace AIT.TFS.SyncService.Contracts.Enums
{
    /// <summary>
    /// Enumeration defines how to find the field in the linked work item definition.
    /// </summary>
    public enum FieldPositionType
    {
        /// <summary>
        /// Field should be invisible.
        /// </summary>
        Hidden,

        /// <summary>
        /// Field is in the first row of numbered list item.
        /// </summary>
        NumberedListItem,

        /// <summary>
        /// Field is at the end of the first row of numbered list item in square brackets. 'Title [First][Second][Third]'
        /// </summary>
        NumberedListItemFirstAddition,

        /// <summary>
        /// Field is at the end of the first row of numbered list item in square brackets. 'Title [First][Second][Third]'
        /// </summary>
        NumberedListItemSecondAddition,

        /// <summary>
        /// Field is at the end of the first row of numbered list item in square brackets. 'Title [First][Second][Third]'
        /// </summary>
        NumberedListItemThirdAddition,

        /// <summary>
        /// Field is remainder of the numbered list item.
        /// </summary>
        Remainder,
    }
}
