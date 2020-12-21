namespace TFS.SyncService.W2TConsole
{
    using System;
    using System.Collections.Generic;
    using AIT.TFS.SyncService.Contracts.Enums;
    using AIT.TFS.SyncService.Contracts.InfoStorage;
    using AIT.TFS.SyncService.Factory;
    using AIT.TFS.SyncService.Model.Console;
    using ManyConsole;
    using Resources = TFS.SyncService.W2TConsole.Properties.Resources;

    public class WordToTFS2013
    {
        public static void Main(string[] args)
        {
            var commands = GetCommands();
            var code = ConsoleCommandDispatcher.DispatchCommand(commands, args, Console.Out);

            var infoService = SyncServiceFactory.GetService<IInfoStorageService>(); //internally initialized by Adapter
            if (infoService != null && infoService.UserInformation.Count > 0)
            {
                foreach (var info in infoService.UserInformation)
                {
                    switch (info.Type)
                    {
                        case UserInformationType.Success:
                            Console.ForegroundColor = ConsoleColor.Green;
                            break;
                        case UserInformationType.Warning:
                        case UserInformationType.Unmodified:
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            code = -2;
                            break;
                        case UserInformationType.Error:
                        case UserInformationType.Conflicting:
                        default:
                            Console.ForegroundColor = ConsoleColor.Red;
                            code = -2;
                            break;
                    }
                    Console.WriteLine(info.ToString());
                    Console.ResetColor();
                }
            }
            HandleReturnCode(code);
        }

        /// <summary>
        /// Obtain all possible commands for many console
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<ConsoleCommand> GetCommands()
        {
            return ConsoleCommandDispatcher.FindCommandsInSameAssemblyAs(typeof(WordToTFS2013));
        }

        /// <summary>
        /// Return codes
        /// 0 Everything alright
        /// -1 Code from ManyConsole
        /// -2 An error occured during creation
        /// </summary>
        /// <param name="code"></param>
        private static void HandleReturnCode(int code)
        {
            switch (code)
            {
                case 0:
                    ConsoleExtensionLogging.LogMessage(Resources.ReportCreationSuccessfull, ConsoleExtensionLogging.LogLevel.Both);
                    Environment.Exit(0);
                    break;
                case -2:
                    ConsoleExtensionLogging.LogMessage(Resources.ReportCreationFailed, ConsoleExtensionLogging.LogLevel.Both);
                    Environment.Exit(1);
                    break;
             }
        }
    }
}
