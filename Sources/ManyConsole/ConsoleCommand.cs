﻿namespace ManyConsole
{
    #region Usings
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Internal;
    using NDesk.Options;
    using AIT.TFS.SyncService.Model.TemplateManagement;
    #endregion


    public abstract class ConsoleCommand : ConsoleUtil
    {
        #region Constructor
        public ConsoleCommand()
        {
            OneLineDescription = "";
            Options = new OptionSet();
            TraceCommandAfterParse = true;
            RemainingArgumentsCount = 0;
            RemainingArgumentsHelpText = "";
            OptionsHasd = new OptionSet();
            RequiredOptions = new List<RequiredOptionRecord>();

            // The initiation of the template manager updates the templates
            // ReSharper disable UnusedVariable
            var templateManager = new TemplateManager();
            // ReSharper restore UnusedVariable
        }
        #endregion

        #region Properties
        public string Command { get; private set; }
        public string OneLineDescription { get; private set; }
        public OptionSet Options { get; protected set; }
        public bool TraceCommandAfterParse { get; private set; }
        public int? RemainingArgumentsCount { get; private set; }
        public string RemainingArgumentsHelpText { get; private set; }
        private OptionSet OptionsHasd { get; set; }
        private List<RequiredOptionRecord> RequiredOptions { get; set; }
        #endregion

        #region Public methods

        public ConsoleCommand IsCommand(string command, string oneLineDescription = "")
        {
            Command = command;
            OneLineDescription = oneLineDescription;
            return this;
        }

        public ConsoleCommand HasAdditionalArguments(int? count = 0, string helpText = "")
        {
            RemainingArgumentsCount = count;
            RemainingArgumentsHelpText = helpText;
            return this;
        }

        public ConsoleCommand AllowsAnyAdditionalArguments(string helpText = "")
        {
            HasAdditionalArguments(null, helpText);
            return this;
        }

        public ConsoleCommand SkipsCommandSummaryBeforeRunning()
        {
            TraceCommandAfterParse = false;
            return this;
        }

        public ConsoleCommand HasOption(string prototype, string description, Action<string> action)
        {
            OptionsHasd.Add(prototype, description, action);

            return this;
        }

        public ConsoleCommand HasRequiredOption(string prototype, string description, Action<string> action)
        {
            HasRequiredOption<string>(prototype, description, action);

            return this;
        }

        public ConsoleCommand HasOption<T>(string prototype, string description, Action<T> action)
        {
            OptionsHasd.Add(prototype, description, action);
            return this;
        }

        public ConsoleCommand HasRequiredOption<T>(string prototype, string description, Action<T> action)
        {
            var requiredRecord = new RequiredOptionRecord();

            var previousOptions = OptionsHasd.ToArray();

            OptionsHasd.Add<T>(prototype, description, s =>
            {
                requiredRecord.WasIncluded = true;
                action(s);
            });

            var newOption = OptionsHasd.Single(o => !previousOptions.Contains(o));

            requiredRecord.Name = newOption.GetNames().OrderByDescending(n => n.Length).First();

            RequiredOptions.Add(requiredRecord);

            return this;
        }

        public ConsoleCommand HasOption(string prototype, string description, OptionAction<string, string> action)
        {
            OptionsHasd.Add(prototype, description, action);
            return this;
        }

        public ConsoleCommand HasOption<TKey, TValue>(string prototype, string description, OptionAction<TKey, TValue> action)
        {
            OptionsHasd.Add(prototype, description, action);
            return this;
        }

        public virtual void CheckRequiredArguments()
        {
            var missingOptions = this.RequiredOptions
                .Where(o => !o.WasIncluded).Select(o => o.Name).OrderBy(n => n).ToArray();

            if (missingOptions.Any())
            {
                throw new ConsoleHelpAsException("Missing option: " + String.Join(", ", missingOptions));
            }
        }

        public virtual int? OverrideAfterHandlingArgumentsBeforeRun(string[] remainingArguments)
        {
            return null;
        }

        public abstract int Run(string[] remainingArguments);

        public OptionSet GetActualOptions()
        {
            var result = new OptionSet();

            foreach (var option in Options)
                result.Add(option);

            foreach (var option in OptionsHasd)
                result.Add(option);

            return result;
        }
        #endregion
    }
}