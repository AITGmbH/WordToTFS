using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.TfsHelper;
using AIT.TFS.SyncService.Factory;
using Microsoft.TeamFoundation.TestManagement.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects
{
    using Common.Helper;
    using Properties;

    /// <summary>
    /// Implementation of the TFS test steps field. Assigning a value to this field will temporarily save changes
    /// and set a work item flag that triggers generating and saving test steps when the work item is saved.
    /// </summary>
    internal class TfsTestCaseField : TfsField
    {
        private string _value;

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTestCaseField"/> class.
        /// </summary>
        /// <param name="workItem">Work item to present one field from.</param>
        /// <param name="configurationFieldItem">Configuration field item.</param>
        internal TfsTestCaseField(TfsWorkItem workItem, IConfigurationFieldItem configurationFieldItem)
            : base(workItem, configurationFieldItem)
        {
            //TODO MIS 25.06.2015 Commented out because currently appeared in build. Recheck
            //Debug.Assert(ReferenceName.Equals(FieldReferenceNames.TestSteps), "Only use this field implementation for TestSteps");
            //Debug.Assert(Configuration.FieldValueType == FieldValueType.BasedOnFieldType, "make sure this field is correctly set up");
        }
        #endregion

        #region IField Members

        /// <summary>
        /// The value of the current field.
        /// </summary>
        public override string Value
        {
            get
            {
                // If the field is dirty, return the temporary change instead the actual TFS values
                if (_value != null) return _value;

                object value = WorkItem.WorkItem.Fields[ReferenceName].Value;
                if (null != value)
                {
                    return GetTestCaseValue();
                }

                return string.Empty;
            }
            set
            {
                _value = value;
            }
        }

        /// <summary>
        /// Gets special value which needs special handling.
        /// </summary>
        /// <value> Returns special value data or information about this data. </value>
        public override object MicroDocument
        {
            get { return null; }
            set { }
        }

        /// <summary>
        /// Compares the value of this field to the value of another field.
        /// </summary>
        /// <param name="value">The value to compare to.</param>
        /// <param name="ignoreFormatting">Sets whether formatting is ignored when comparing html fields</param>
        /// <returns>True if the values are equal, False if not.</returns>
        public override bool CompareValue(string value, bool ignoreFormatting)
        {
            return Value == value;
        }

        #endregion

        /// <summary>
        /// Gets Test Case value from Step actions.
        /// </summary>
        /// <returns></returns>
        private string GetTestCaseValue()
        {
            //If work item is a shared step
            if (WorkItem.TestCase == null && WorkItem.WorkItem.Fields.Contains(FieldReferenceNames.TestSteps))
            {
                var tfs = WorkItem.WorkItemStore.TeamProjectCollection; //TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverName));
                var testService = tfs.GetService<ITestManagementService>();
                var project = testService.GetTeamProject(WorkItem.WorkItem.Project.Name);

                // Query the Shared Step by id
                const string interestedFields = "[System.Id], [System.Title]";
                var query = $"SELECT {interestedFields} FROM WorkItems WHERE System.Id = '{WorkItem.Id}'";
                var sharedSteps = project.SharedSteps.Query(query);

                if (sharedSteps.Count() == 1)
                {
                    var foundSharedSteps = sharedSteps.First();

                    var step = string.Empty;
                    foreach (ITestAction t in foundSharedSteps.Actions)
                    {
                        var testStep = t as ITestStep;
                        if (testStep != null)
                        {
                            step += HtmlHelper.ExtractPlaintext(testStep.Title);

                            if (!string.IsNullOrEmpty(testStep.ExpectedResult))
                            {
                                step += Configuration.TestCaseStepDelimiter + HtmlHelper.ExtractPlaintext(testStep.ExpectedResult);
                            }
                            step += '\n';
                        }
                    }
                    return step.TrimEnd('\n');  // make sure this generated string matches the generated string from word work item. If not, syncservice will always assign the word value and all test cases will always be recreated and saved
                }
            }

            if (WorkItem.TestCase != null && WorkItem.TestCase.Actions.Count > 0)
            {


                var step = string.Empty;
                foreach (ITestAction t in WorkItem.TestCase.Actions)
                {
                    var testStep = t as ITestStep;
                    if (testStep != null)
                    {
                        step += HtmlHelper.ExtractPlaintext(testStep.Title);

                        if (string.IsNullOrEmpty(testStep.ExpectedResult) == false)
                        {
                            step += Configuration.TestCaseStepDelimiter + HtmlHelper.ExtractPlaintext(testStep.ExpectedResult);
                        }

                        step += '\n';
                    }
                }

                return step.TrimEnd('\n');  // make sure this generated string matches the generated string from word work item. If not, syncservice will always assign the word value and all test cases will always be recreated and saved
            }

            return string.Empty;
        }

        /// <summary>
        /// Sets Test Case value to Work Item field.
        /// </summary>
        public bool SaveTestCaseValue()
        {
            if (WorkItem.WorkItem.IsDirty)
            {
                SyncServiceTrace.W(Resources.CannotSaveTestSteps, WorkItem.Id);
                return false;
            }

            if (!IsEditable) return false;
            if (_value == null) return true;

            if (WorkItem.TestCase == null && WorkItem.WorkItem.Fields.Contains(FieldReferenceNames.TestSteps))
            {
                throw new NotImplementedException("Shared Steps cannot yet be published to TFS");
            }

            if (WorkItem.TestCase != null)
            {
                // make sure we work with the latest version or we run into
                // an "already saved conflict" when trying to save test steps
                WorkItem.TestCase.Refresh();
                WorkItem.TestCase.Actions.Clear();

                var steps = _value.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                _value = null;
                foreach (var step in steps)
                {
                    var testStep = WorkItem.TestCase.CreateTestStep();
                    if (testStep != null)
                    {
                        var components = step.Split(new[] { Configuration.TestCaseStepDelimiter }, StringSplitOptions.None);
                        if (components.Length > 0)
                        {
                            testStep.Title = new ParameterizedString(components[0]);
                            if (components.Length > 1)
                            {
                                testStep.ExpectedResult = new ParameterizedString(components[1]);
                            }
                            WorkItem.TestCase.Actions.Add(testStep);
                        }
                        else
                        {
                            SyncServiceTrace.E(Resources.ParsingTestStepFailed, step, Configuration.TestCaseStepDelimiter);
                        }
                    }
                }

                WorkItem.TestCase.Save();

                // make sure items with only changes to the test steps are not cached
                // (and therefore appear as skipped rather then published successfully)
                WorkItem.Refresh();
            }
            return true;
        }
    }
}
