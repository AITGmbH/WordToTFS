#region Usings
using System;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.Model.Console;
using TFS.SyncService.W2TConsole.Enums;
using TFS.SyncService.W2TConsole.Properties;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Factory;
#endregion

namespace TFS.SyncService.W2TConsole.Commands
{
    /// <summary>
    /// Command responsible to create a test specification report.
    /// </summary>
    internal class TestSpecReport : WordToTFSCommand
    {
        /// <summary>
        /// Create a TestSpecificationReport
        /// </summary>
        public TestSpecReport()
        {
            //Command Name
            IsCommand("TestSpecReport", "Create a test specification report");
        }

        /// <summary>
        /// Execution of the console parameter
        /// </summary>
        /// <param name="remainingArguments"></param>
        /// <returns></returns>
        public override int Run(string[] remainingArguments)
        {
            try
            {
                var conf = CreateDocumentConfig();

                if (conf == null || conf.TestSpecSettings == null)
                {
                    throw new ArgumentException(string.Format(Resources.ConfigurationFileMisconfigured, "TestSpecificationConfiguration"));                    
                }

                // Check if all required test specification settings are given and valid
                var testSpecSettings = conf.TestSpecSettings;
                if (testSpecSettings.TestPlan == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsTestPlanError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsTestPlanError);
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsErrorInstruction, ConsoleExtensionLogging.LogLevel.Console);
                    return CommandReturnCodeFail;
                }

                if (testSpecSettings.CreateDocumentStructure && testSpecSettings.DocumentStructure == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsDocumentStructureError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsDocumentStructureError);
                    return CommandReturnCodeFail;
                }

                DocumentStructureType structure;
                if (testSpecSettings.CreateDocumentStructure && !Enum.TryParse(testSpecSettings.DocumentStructure, out structure))
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsInvalidDocumentStructureError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsInvalidDocumentStructureError);
                    return CommandReturnCodeFail;
                }
                if (testSpecSettings.IncludeTestConfigurations && testSpecSettings.TestConfigurationsPosition == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsConfigurationPositionError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsConfigurationPositionError);
                    return CommandReturnCodeFail;
                }

                TestConfigurationPosition position;
                if (testSpecSettings.IncludeTestConfigurations && !Enum.TryParse(testSpecSettings.TestConfigurationsPosition, out position))
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsInvalidConfigurationPositionError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsInvalidConfigurationPositionError);
                    return CommandReturnCodeFail;
                }

                if (testSpecSettings.SortTestCasesBy == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsTestCasesSortError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsTestCasesSortError);
                    return CommandReturnCodeFail;
                }

                TestCaseSortType sort;
                if (testSpecSettings.SortTestCasesBy == null && Enum.TryParse(testSpecSettings.SortTestCasesBy, out sort))
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsInvalidTestCasesSortError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsInvalidTestCasesSortError);
                    return CommandReturnCodeFail;
                }

                var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
                crm.CreateTestSpecDocument(conf);
            }
            catch (Exception e)
            {
                ConsoleExtensionLogging.LogMessage(e.Message, ConsoleExtensionLogging.LogLevel.Both);
                SyncServiceTrace.D(e.Message);
                return CommandReturnCodeFail;
            }
            return CommandReturnCodeSuccess;
        }
    }
}
