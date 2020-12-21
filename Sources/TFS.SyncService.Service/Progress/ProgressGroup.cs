namespace AIT.TFS.SyncService.Service.Progress
{
    /// <summary>
    /// Class is a container for one progress group.
    /// </summary>
    internal class ProgressGroup
    {
        /// <summary>
        /// Gets or sets the count of ticks in this progress group.
        /// </summary>
        public int CountOfTicks { get; set; }

        /// <summary>
        /// Gets or sets the actual tick number.
        /// </summary>
        public int ActualTick { get; set; }

        /// <summary>
        /// Gets or sets the actual progress text.
        /// </summary>
        public string Text { get; set; }
    }
}