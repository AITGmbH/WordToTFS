namespace AIT.TFS.SyncService.Contracts.Enums.Model
{
    /// <summary>
    /// Enumeration defines all possible position of the configurations in the report
    /// </summary>
    public enum ConfigurationPositionType
    {
        /// <summary>
        /// The configuration part is printed once above the testplan
        /// </summary>
        AboveTestPlan,

        /// <summary>
        /// The configuration part is printed beneath the testplan
        ///  </summary>
        BeneathTestPlan,

        /// <summary>
        /// The configruation part is printed beaneath all TestSuitesinfomration
        /// </summary>
        BeneathTestSuites,

        /// <summary>
        /// The configuration part is printed beanth the first TestSuiteInformation
        /// </summary>
        BeneathFirstTestSuite
    }
}