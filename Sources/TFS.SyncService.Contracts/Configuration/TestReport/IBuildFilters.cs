using System.Collections.Generic;
namespace AIT.TFS.SyncService.Contracts.Configuration.TestReport
{
    /// <summary>
    ///  Interface defines build filters for test result report.
    /// </summary>
    public interface IBuildFilters
    {
        /// <summary>
        /// Gets the list of configured build qualities to use as filter for builds.
        /// </summary>
        List<string> BuildQualities { get; set; }

        /// <summary>
        /// Gets the list of configured build names to use as filter for builds.
        /// </summary>
        List<string> BuildNames { get; set; }

        /// <summary>
        /// Gets the list of configured build ages to use as filter for builds.
        /// </summary>
        string BuildAge { get; set; }

        /// <summary>
        /// Gets the list of configured build tags to use as filter for builds.
        /// </summary>
        List<string> BuildTags { get; set; }
    }
}
