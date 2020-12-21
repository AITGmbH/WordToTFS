namespace TFS.SyncService.W2TConsole.Enums
{
    /// <summary>
    /// Enumeration defines, which document structure can be used for test reports
    /// </summary>
    public enum TestConfigurationPosition
    {
        /// <summary>
        /// Create configuration part above test plan
        /// </summary>
        AboveTestPlan,

        /// <summary>
        /// Create configuration part beneath test plan
        /// </summary>
        BeneathTestPlan,

        /// <summary>
        /// Create configuration part beneath test suites
        /// </summary>
        BeneathTestSuites,

        /// <summary>
        /// Create configuration part beneath first test suite
        /// </summary>
        BeneathFirstTestSuite
    }
}
