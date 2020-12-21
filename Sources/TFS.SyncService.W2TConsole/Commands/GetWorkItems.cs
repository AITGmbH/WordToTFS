#region Usings
using System;
using AIT.TFS.SyncService.Model.Console;
using TFS.SyncService.W2TConsole.Properties;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Factory;
using System.Diagnostics;
#endregion

namespace TFS.SyncService.W2TConsole.Commands
{
    internal class GetWorkItems : WordToTFSCommand
    {
        private string _workitems;
        private string _query;

        /// <summary>
        /// Get the specified workitem
        /// Has two Options - Per Id oder per query - one of these options has to be chosen.
        /// </summary>
        public GetWorkItems()
        {
            //Command Name
            IsCommand("GetWorkItems", "Create a document containing different workItems");

            //Select workitems per id
            HasOption("w|WorkitemIDs=", "Get workitems with the specified IDs separated by \",\"", w => _workitems = w);
            //Select workitems per query
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
                // One of the two options has to be chosen
                if (_workitems == null && _query == null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.OptionsNullError, ConsoleExtensionLogging.LogLevel.Both);
                    ConsoleExtensionLogging.LogMessage(Resources.OptionsNullErrorInstruction, ConsoleExtensionLogging.LogLevel.Console);
                    SyncServiceTrace.D(Resources.OptionsNullError);
                    return CommandReturnCodeFail;
                }
                // It isn't allowed to choose both options simultaneously
                if(_workitems != null && _query != null)
                {
                    ConsoleExtensionLogging.LogMessage(Resources.OptionsOverloadedError, ConsoleExtensionLogging.LogLevel.Both);
                    SyncServiceTrace.D(Resources.OptionsOverloadedError);
                    return CommandReturnCodeFail;
                }

                var createDocumentModel = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
                if (_workitems != null)
                {
                      createDocumentModel.CreateWorkItemDocument(conf, _workitems, "ByID");
                }
                else
                {
                      createDocumentModel.CreateWorkItemDocument(conf, _query, "ByQuery");
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
