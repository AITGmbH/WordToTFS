#region Usings
using System;
using AIT.TFS.SyncService.Service.Configuration.Serialization.Console;
using ManyConsole;
using System.Diagnostics;
using AIT.TFS.SyncService.Factory;
#endregion

namespace TFS.SyncService.W2TConsole.Commands
{
    internal abstract class WordToTFSCommand : ConsoleCommand
    {
        /// Optional Options - but either _configFile or all other options can be used - can't be used together!
        private string _server;
        private string _project;
        private string _filename;
        private string _templateName;
        private string _templateFile;
        private bool _overwrite;
        private bool _closeOnFinish;
        private string _configFile;
        private bool _wordHidden;
        private int _debugLevel;
        protected const int CommandReturnCodeFail = -2;
        protected const int CommandReturnCodeSuccess = 0;

        internal WordToTFSCommand()
        {
            IsCommand("Settings", "Settings");

            HasOption("o|Overwrite=", "Overwrite existing files.", o => _overwrite = true);
            HasOption("c|Close=", "Close after finish.", c => _closeOnFinish = true);
            HasOption("s|ServerName=", "The URL of the TFS/VSTS server (e.g. http://example.com:8080/tfs/YourCollection or https://tfs.example.com/tfs/YourCollection).", s => _server = s);
            HasOption("p|Project=", "The name of the project.", p => _project = p);
            HasOption("f|FileName=", "the target file where the document should be saved to.", f => _filename = f);
            HasOption("t|TemplateName=", "The templateName which should be used for the creation of the document.", t => _templateName = t);
            HasOption("h|WordHidden=", "Determines whether the word document is opened during report creation.", h => _wordHidden = Convert.ToBoolean(h));
            HasOption("d|DotxTemplate=", "The path to a dotx file which should be used for the creation of the document.", t => _templateFile = t);
            // If this option is used, the other options can't be specified
            HasOption("co|ConfigFile=", "The path to to the XML-ConfigFile, which can be used to specify the settings.", co => _configFile = co);
            HasOption("l|DebugLevel=", "The level which should be used for logging.", l => _debugLevel = Convert.ToInt32(l));
        }

        public DocumentConfiguration CreateDocumentConfig()
        {
            try
            {
                var documentConfiguration = new DocumentConfiguration(_server, _project, _filename, _templateName, _overwrite, _closeOnFinish, _configFile, _wordHidden, _templateFile, (TraceLevel)_debugLevel);
                SyncServiceTrace.DebugLevel = new TraceSwitch("DebugLevel", "The Output level of tracing", documentConfiguration.Settings.DebugLevel.ToString());
                return documentConfiguration;
            }
            catch (Exception e)
            {
                // ReSharper disable once PossibleIntendedRethrow
                throw e;
            }
        }
    }
}
