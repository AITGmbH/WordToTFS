namespace AIT.TFS.SyncService.Contracts.Configuration
{
    /// <summary>
    /// Defines modes on how to deal with inline shapes that word does not export without text around it.
    /// </summary>
    public enum ShapeOnlyWorkaroundMode
    {
        /// <summary>
        /// If there is a field with only a shape an no text, try to add a space at the end.
        /// </summary>
        AddSpace,

        /// <summary>
        /// If the is a field with only a shape and no text, raise an error
        /// </summary>
        ShowAsError
            
    }
}
