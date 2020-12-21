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
    internal class TestResultReport : WordToTFSCommand
    {
        /// <summary>
        /// Create a TestSpecificationReport
        /// </summary>
        public TestResultReport()
        {
            //Command Name
            IsCommand("TestResultReport", "Create a test result report");
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
                // Check if all required test specification settings are given and valid

                if (conf == null || conf.TestResultSettings == null)
                {
                    throw new ArgumentException(string.Format(Resources.ConfigurationFileMisconfigured, "TestResultConfiguration"));
                }

                var testResultSettings = conf.TestResultSettings;
                if (testResultSettings.TestPlan == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsTestPlanError, ConsoleExtensionLogging.LogLevel.Both);
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsErrorInstruction, ConsoleExtensionLogging.LogLevel.Console);
                    SyncServiceTrace.D(Resources.TestSettingsTestPlanError);
                    return CommandReturnCodeFail;
                }

                if (testResultSettings.CreateDocumentStructure && testResultSettings.DocumentStructure == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsDocumentStructureError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsDocumentStructureError);
                    return CommandReturnCodeFail;
                }

                DocumentStructureType structure;
                if (testResultSettings.CreateDocumentStructure && !Enum.TryParse(testResultSettings.DocumentStructure, out structure))
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsInvalidDocumentStructureError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsInvalidDocumentStructureError);
                    return CommandReturnCodeFail;
                }

                if (testResultSettings.IncludeTestConfigurations && testResultSettings.TestConfigurationsPosition == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsConfigurationPositionError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsConfigurationPositionError);
                    return CommandReturnCodeFail;
                }

                TestConfigurationPosition position;
                if (testResultSettings.IncludeTestConfigurations && !Enum.TryParse(testResultSettings.TestConfigurationsPosition, out position))
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsInvalidConfigurationPositionError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsConfigurationPositionError);
                    return CommandReturnCodeFail;
                }

                if (testResultSettings.SortTestCasesBy == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsTestCasesSortError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsConfigurationPositionError);
                    return CommandReturnCodeFail;
                }

                TestCaseSortType sort;
                if (testResultSettings.SortTestCasesBy == null && Enum.TryParse(testResultSettings.SortTestCasesBy, out sort))
                {
                    ConsoleExtensionLogging.LogMessage(Resources.TestSettingsInvalidTestCasesSortError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.TestSettingsConfigurationPositionError);
                    return CommandReturnCodeFail;
                }

                var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
                crm.CreateTestResultDocument(conf);
            }
            catch (Exception e)
            {
                ConsoleExtensionLogging.LogMessage(e.Message, ConsoleExtensionLogging.LogLevel.Both);
                SyncServiceTrace.D(e.Message);
                return CommandReturnCodeFail;
            }

            return 0;
        }
    }
}
