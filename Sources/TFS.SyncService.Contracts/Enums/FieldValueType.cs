namespace AIT.TFS.SyncService.Contracts.Enums
{
    /// <summary>
    /// Defines the type of one field in work item.
    /// </summary>
    public enum FieldValueType
    {
        /// <summary>
        /// Field is plain text
        /// </summary>
        PlainText,

        /// <summary>
        /// Field is html formatted
        /// </summary>
        HTML,

        /// <summary>
        /// Field handle its value based on field type.
        /// </summary>
        BasedOnFieldType,

        /// <summary>
        /// Field values are presented as drop down list
        /// </summary>
        DropDownList,

        /// <summary>
        /// Field gets its value from a variable defined in the WordToTfs template.
        /// </summary>
        BasedOnVariable,

        /// <summary>
        /// Field gets its value from a system variable.
        /// </summary>
        BasedOnSystemVariable
    }

    public static class FieldValueTypeExtension
    {
        public static bool IsVariable(this FieldValueType instance)
        {
            return instance == FieldValueType.BasedOnSystemVariable || instance == FieldValueType.BasedOnVariable;
        }
    }
}