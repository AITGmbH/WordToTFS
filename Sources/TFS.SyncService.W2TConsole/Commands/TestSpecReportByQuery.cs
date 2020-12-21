#region Usings
using System;
using AIT.TFS.SyncService.Model.Console;
using TFS.SyncService.W2TConsole.Properties;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Factory;
#endregion

namespace TFS.SyncService.W2TConsole.Commands
{
    internal class TestSpecReportByQuery : WordToTFSCommand
    {
        private string _workitems;
        private string _query;

        /// <summary>
        /// Create a TestSpecificationReport
        /// Has three Options - Per Id, per query or per title - one of these options has to be chosen.
        /// </summary>
        public TestSpecReportByQuery()
        {
            // Command Name
            IsCommand("TestSpecReportByQuery", "Create a test specification report by query");

            // Select workitems per id
            HasOption("w|WorkitemIDs=", "Get workitems with the specified IDs separated by \",\"", w => _workitems = w);
            // Select workitems per query
            HasOption("q|WorkitemQuery=", "Get workitems from the specified work item query", q => _query = q);
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

                // One of the two options has to be chosen
                if (_workitems == null && _query == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.OptionsNullError, ConsoleExtensionLogging.LogLevel.Both);
                    ConsoleExtensionLogging.LogMessage(Resources.OptionsNullErrorInstruction, ConsoleExtensionLogging.LogLevel.Console);
                    SyncServiceTrace.D(Resources.OptionsNullError);
                    return CommandReturnCodeFail;
                }

                // It isn't allowed to choose both options simultaneously
                if (_workitems != null && _query != null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.OptionsOverloadedError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.OptionsOverloadedError);
                    return CommandReturnCodeFail;
                }

                if (_workitems != null)
                {
                    var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
                    crm.CreateTestSpecDocumentByQuery(conf, _workitems, "ByID");
                }
                // Get work items by query
                else
                {
                    var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
                    crm.CreateTestSpecDocumentByQuery(conf, _query, "ByQuery");
                }
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
