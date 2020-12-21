#region Usings
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Contracts.TfsHelper;
using AIT.TFS.SyncService.Contracts.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Service.InfoStorage;
using AIT.TFS.SyncService.Service.Properties;
using AIT.TFS.SyncService.Service.Utils;
using AIT.TFS.SyncService.Service.WorkItemObjects;
using Microsoft.TeamFoundation.Common;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Threading;
#endregion

namespace AIT.TFS.SyncService.Service
{
    /// <summary>
    /// Class implements the functionality of the service for the synchronization of two adapters.
    /// </summary>
    internal class WorkItemSyncService : IWorkItemSyncService
    {
        #region Fields
        private readonly ISystemVariableService _systemVariableService;
        private IConfiguration _configuration;
        private ProjectPathConverter _pathConverter;
        #endregion

        #region Constructor
        public WorkItemSyncService(ISystemVariableService systemVariableService)
        {
            _systemVariableService = systemVariableService;
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Import work items from TFS into Word.
        /// </summary>
        /// <param name="sourceTfs">Adapter to the TFS server.</param>
        /// <param name="destinationWord">Adapter to Word document.</param>
        /// <param name="importWorkItems">List of work items from the source adapter to import.</param>
        /// <param name="configuration">The configuration to be used during the refresh operation</param>
        public void Refresh(IWorkItemSyncAdapter sourceTfs, IWorkItemSyncAdapter destinationWord, IEnumerable<IWorkItem> importWorkItems, IConfiguration configuration)
        {
            if (sourceTfs == null) throw new ArgumentNullException("sourceTfs");
            if (destinationWord == null) throw new ArgumentNullException("destinationWord");
            if (importWorkItems == null) throw new ArgumentNullException("importWorkItems");
            if (configuration == null) throw new ArgumentNullException("configuration");

            _configuration = configuration;
            _pathConverter = new ProjectPathConverter(((ITfsService)sourceTfs).ProjectName);
            var workItems = importWorkItems.ToList();

            SyncServiceTrace.I(Resources.RefreshInfo, workItems.Count(), string.Join(",", workItems.Select(x => x.Id.ToString(CultureInfo.InvariantCulture)).ToArray()));

            var progressService = SyncServiceFactory.GetService<IProgressService>();

            // step 1: open both adapters
            progressService.EnterProgressGroup(2, Resources.GetWorkItems_Importing);
            if (!sourceTfs.Open(workItems.Select(x => x.Id).ToArray()) || !destinationWord.Open(null)) return;
            progressService.DoTick();

            // step 2: refreshing / importing all items
            var mappings = sourceTfs.WorkItems.ToDictionary<IWorkItem, IWorkItem, IWorkItem>(workItem => workItem, workItem => null);
            ReverseSync(mappings, sourceTfs, destinationWord);
            progressService.LeaveProgressGroup();
        }

        /// <summary>
        /// Imports work work items from TFS into Word and refreshes existing work items.
        /// </summary>
        /// <param name="sourceTfs">Source adapter represented by TFS server.</param>
        /// <param name="destinationWord">Destination adapter represented by Word document.</param>
        /// <param name="importWorkItems">List of work items to import.</param>
        /// <param name="configuration">Configuration to use during refresh</param>
        /// <param name="testReportHelper">The test report helper.</param>
        /// <param name="testCases">The test cases.</param>
        public void RefreshAndSubstituteTestItems(IWorkItemSyncAdapter sourceTfs,
            IWorkItemSyncAdapter destinationWord,
            IEnumerable<IWorkItem> importWorkItems,
            IConfiguration configuration,
            ITestReportHelper testReportHelper,
            IDictionary<int, ITfsTestCaseDetail> testCases)
        {
            if (sourceTfs == null) throw new ArgumentNullException("sourceTfs");
            if (destinationWord == null) throw new ArgumentNullException("destinationWord");
            if (importWorkItems == null) throw new ArgumentNullException("importWorkItems");
            if (configuration == null) throw new ArgumentNullException("configuration");
            if (testReportHelper == null) throw new ArgumentNullException("testReportHelper");
            if (testCases == null) throw new ArgumentNullException("testCases");

            var testCaseElementTemplateName = configuration.ConfigurationTest.ConfigurationTestSpecification.TestCaseElementTemplate;
            var testCaseHeaderTemplateName = configuration.ConfigurationTest.GetHeaderTemplate(testCaseElementTemplateName);

            _configuration = configuration;
            _pathConverter = new ProjectPathConverter(((ITfsService)sourceTfs).ProjectName);
            var workItems = importWorkItems.ToList();

            SyncServiceTrace.I(Resources.RefreshInfo, workItems.Count(), string.Join(",", workItems.Select(x => x.Id.ToString(CultureInfo.InvariantCulture)).ToArray()));
            var progressService = SyncServiceFactory.GetService<IProgressService>();

            // step 1: open both adapters
            progressService.EnterProgressGroup(2, Resources.GetWorkItems_Importing);
            if (!sourceTfs.Open(workItems.Select(x => x.Id).ToArray()) || !destinationWord.Open(null)) return;
            progressService.DoTick();

            // step 2: refreshing / importing all items
            var mappings = sourceTfs.WorkItems.ToDictionary<IWorkItem, IWorkItem, IWorkItem>(workItem => workItem, workItem => null);

            progressService.EnterProgressGroup(2);
            progressService.EnterProgressGroup(mappings.Count);

            // Create threads to refresh work items in parallel
            const int ThreadCount = 4;
            var finished = new SynchronizedCollection<IWorkItem>();

            RefreshWorkItemsInParallel(ThreadCount, mappings, finished);

            var linkMapping = new Dictionary<IWorkItem, IWorkItem>();

            if (WriteWorkItemsToDocument(sourceTfs, destinationWord, testReportHelper, testCases, mappings, finished, testCaseElementTemplateName, testCaseHeaderTemplateName, linkMapping, progressService))
                return;
            progressService.LeaveProgressGroup();
            progressService.DoTick();
            LinkWorkItems(linkMapping, sourceTfs, Direction.TfsToOther);
            progressService.LeaveProgressGroup();

            progressService.LeaveProgressGroup();
        }

        /// <summary>
        /// Imports work work items from TFS into Word and refreshes existing work items.
        /// </summary>
        /// <param name="sourceTfs">Source adapter represented by TFS server.</param>
        /// <param name="destinationWord">Destination adapter represented by Word document.</param>
        /// <param name="importWorkItems">List of work items to import.</param>
        /// <param name="configuration">Configuration to use during refresh</param>
        /// <param name="testReportHelper">The test report helper.</param>
        /// <param name="sharedSteps">The shared steps.</param>
        public void RefreshAndSubstituteSharedStepItems(IWorkItemSyncAdapter sourceTfs,
            IWorkItemSyncAdapter destinationWord,
            IEnumerable<IWorkItem> importWorkItems,
            IConfiguration configuration,
            ITestReportHelper testReportHelper,
            IDictionary<int, ITfsSharedStepDetail> sharedSteps)
        {
            if (sourceTfs == null) throw new ArgumentNullException("sourceTfs");
            if (destinationWord == null) throw new ArgumentNullException("destinationWord");
            if (importWorkItems == null) throw new ArgumentNullException("importWorkItems");
            if (configuration == null) throw new ArgumentNullException("configuration");
            if (testReportHelper == null) throw new ArgumentNullException("testReportHelper");
            if (sharedSteps == null) throw new ArgumentNullException("sharedSteps");

            var sharedStepElementTemplateName = configuration.ConfigurationTest.ConfigurationTestSpecification.SharedStepsElementTemplate;
            var testCaseHeaderTemplateName = configuration.ConfigurationTest.GetHeaderTemplate(sharedStepElementTemplateName);

            _configuration = configuration;
            _pathConverter = new ProjectPathConverter(((ITfsService)sourceTfs).ProjectName);
            var workItems = importWorkItems.ToList();

            SyncServiceTrace.I(Resources.RefreshInfo, workItems.Count(), string.Join(",", workItems.Select(x => x.Id.ToString(CultureInfo.InvariantCulture)).ToArray()));
            var progressService = SyncServiceFactory.GetService<IProgressService>();

            // step 1: open both adapters
            progressService.EnterProgressGroup(2, Resources.GetWorkItems_Importing);
            if (!sourceTfs.Open(workItems.Select(x => x.Id).ToArray()) || !destinationWord.Open(null)) return;
            progressService.DoTick();

            // step 2: refreshing / importing all items
            var mappings = sourceTfs.WorkItems.ToDictionary<IWorkItem, IWorkItem, IWorkItem>(workItem => workItem, workItem => null);

            progressService.EnterProgressGroup(2);
            progressService.EnterProgressGroup(mappings.Count);

            // Create threads to refresh work items in parallel
            const int ThreadCount = 4;
            var finished = new SynchronizedCollection<IWorkItem>();

            RefreshWorkItemsInParallel(ThreadCount, mappings, finished);

            var linkMapping = new Dictionary<IWorkItem, IWorkItem>();

            if (WriteWorkItemsToDocument(sourceTfs, destinationWord, testReportHelper, sharedSteps, mappings, finished, sharedStepElementTemplateName, testCaseHeaderTemplateName, linkMapping, progressService))
                return;
            progressService.LeaveProgressGroup();
            progressService.DoTick();
            LinkWorkItems(linkMapping, sourceTfs, Direction.TfsToOther);
            progressService.LeaveProgressGroup();

            progressService.LeaveProgressGroup();
        }


        /// <summary>
        /// Method synchronizes two work adapters.
        /// </summary>
        /// <param name="source">The source of synchronization - synchronize from this work item.</param>
        /// <param name="destination">The destination of synchronization - synchronize to this work item.</param>
        /// <param name="workItems">A list of work items to publish</param>
        /// <param name="forceOverwrite">Sets whether work items are synchronized, even if the destination is more recent</param>
        /// <param name="configuration">Configuration to use during publish</param>
        public void Publish(IWorkItemSyncAdapter source, IWorkItemSyncAdapter destination, IEnumerable<IWorkItem> workItems, bool forceOverwrite, IConfiguration configuration)
        {
            // Materialize work items (a Linq statement), since they will be used often below in various functions,
            // including other Linq statements (then the first Linq statement has to be evaluated again and again ...).
            workItems = workItems.ToList();

            SyncServiceTrace.I(Resources.LogService_Sync_Init);
            SyncServiceTrace.D(Resources.Overwrite, forceOverwrite);
            SyncServiceTrace.D(Resources.PublishWorkItems, workItems.Count(), string.Join(",", workItems.Select(x => x.Id.ToString(CultureInfo.InvariantCulture)).ToArray()));

            if (source == null) throw new ArgumentNullException("source");
            if (destination == null) throw new ArgumentNullException("destination");
            if (configuration == null) throw new ArgumentNullException("configuration");

            _configuration = configuration;
            _pathConverter = new ProjectPathConverter(((ITfsService)destination).ProjectName);

            var progressService = SyncServiceFactory.GetService<IProgressService>();

            // step 1a: open word and validate word items
            progressService.EnterProgressGroup(6, Resources.LO_Publish_01_OpenWord);

            if (!source.Open(null) || progressService.ProgressCanceled)
            {
                goto EXIT;
            }

            // Materialize work items (a Linq statement), since they will be used often below in various functions,
            // including other Linq statements (then the first Linq statement has to be evaluated again and again ...).
            var workItemsToPublish = source.WorkItems.Where(wi => workItems.Any(y => y.IsSameWorkItem(wi))).ToList().AsEnumerable();

            //Set the flag to parse html different
            workItemsToPublish = OverwriteParseHTMLAsPlainTextFlag(workItemsToPublish, workItems);
            var errors = source.ValidateWorkItems();
            if (errors.Count > 0)
            {
                foreach (var error in errors)
                {
                    PublishWorkItemError(error.WorkItem, error.Field, error.Exception.Message);
                }

                goto EXIT;
            }

            // step 2a: open TFS
            progressService.DoTick(Resources.LO_Publish_02_OpenTFS);
            var sourceWorkItems = workItemsToPublish.Where(wi => wi.Id > 0);
            var configurations = new Dictionary<int, IConfigurationItem>();
            foreach (var item in sourceWorkItems)
            {
                configurations.Add(item.Id, item.Configuration);
            }

            // step 2b: compare the fields of word with the fields of tfs
            // If the verifiery returs true, get a list with the wrong mapped items and show the error dialog, if the users responds with overwrite change the parsing, else exclude the items from the sync
            //CompareTypeOfWorkItems(workItemsToPublish, destination, configuration);

            var openResult = destination.OpenWithConfigurations(configurations);
            if (!openResult || progressService.ProgressCanceled)
            {
                goto EXIT;
            }

            // step 3: sync word to TFS
            progressService.DoTick(Resources.LO_Publish_03_SyncToTFS);
            SyncServiceTrace.I(Resources.LogService_Sync_InitForwardSync);
            var mappedItems = ForwardSync(destination, source, workItemsToPublish, forceOverwrite);

            // save links. Needs to be done after all work items exist
            var dirtyItems = mappedItems.Keys.Where(x => x.IsDirty).ToList();
            if (!SyncValidateWorkItems(destination, source) || !SaveChanges(destination, mappedItems) || progressService.ProgressCanceled)
            {
                goto EXIT;
            }

            LinkWorkItems(mappedItems, source, Direction.OtherToTfs);

            // step 4: validate synced values and save changes to TFS
            progressService.DoTick(Resources.LO_Publish_04_SaveToTFS);
            if (!SyncValidateWorkItems(destination, source) || !SaveChanges(destination, mappedItems) || progressService.ProgressCanceled)
            {
                goto EXIT;
            }

            foreach (var item in mappedItems.Where(item => !item.Key.IsDirty))
            {
                if (dirtyItems.Contains(item.Key))
                {
                    Publish(new UserInformation(item.Key)
                    {
                        Destination = item.Value,
                        Type = UserInformationType.Success,
                        Explanation = Resources.Publish_OK
                    });
                }
                else
                {
                    Publish(new UserInformation(item.Key)
                    {
                        Destination = item.Value,
                        Type = UserInformationType.Unmodified,
                        Explanation = Resources.Publish_Skipped
                    });
                }
            }

            // step 5: sync changed TFS items back to word
            progressService.DoTick(Resources.LO_Publish_05_SyncToWord);
            SyncServiceTrace.I(Resources.LogService_Sync_InitReverseSync);
            ReverseSync(mappedItems, destination, source);
            if (progressService.ProgressCanceled) goto EXIT;

            // step 6: save changes to word
            progressService.DoTick(Resources.LO_Publish_06_SaveToWord);
            //if (!SaveChanges(source, mappedItems) || progressService.ProgressCanceled) goto EXIT;

            // step 7: Process the linked items - linked items come from the special cell of the table.
            // this should not be here - SER
            SyncServiceTrace.I(Resources.LogService_Sync_InitLinkedItems);

            foreach (var baseWorkItem in source.WorkItems)
            {
                var workItemToLink = baseWorkItem as IWorkItemLinkedItems;
                if (workItemToLink == null
                    || workItemToLink.LinkedWorkItems == null
                    || workItemToLink.LinkedWorkItems.Count == 0)
                    continue;
                var destinationFrom = destination.WorkItems.Find(workItemToLink.Id);
                if (destinationFrom == null)
                    continue;
                bool addLinkCalledFirstTime = true;
                // Key in Dictionary is a link type.
                // Value in Dictionary is a collection of work items with link type of key.
                foreach (var keyValuePair in workItemToLink.LinkedWorkItems)
                {
                    foreach (var linkedWorkItem in keyValuePair.Value)
                    {
                        var destionationTo = destination.WorkItems.Find(linkedWorkItem.Id);
                        if (destionationTo == null)
                            continue;
                        destinationFrom.AddLinks(destination, new[] { linkedWorkItem.Id }, keyValuePair.Key, addLinkCalledFirstTime);
                        addLinkCalledFirstTime = false;
                    }
                }
            }

            if (!SaveChanges(destination, mappedItems))
                return;

            EXIT:
            if (progressService.ProgressCanceled)
            {
                SyncServiceTrace.W(Resources.Publish_Canceled);

                Publish(new UserInformation
                {
                    Text = Resources.Publish_Canceled,
                    Explanation = Resources.Publish_Canceled_Desc,
                    Type = UserInformationType.Warning
                });
            }

            progressService.LeaveProgressGroup();
            SyncServiceTrace.I(Resources.LogService_Sync_Finished);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Writes the work items to the document.
        /// </summary>
        /// <param name="sourceTfs">The source TFS.</param>
        /// <param name="destinationWord">The destination word.</param>
        /// <param name="testReportHelper">The test report helper.</param>
        /// <param name="testCases">The test cases.</param>
        /// <param name="mappings">The mappings.</param>
        /// <param name="finished">The finished.</param>
        /// <param name="testCaseElementTemplateName">Name of the test case element template.</param>
        /// <param name="testCaseHeaderTemplateName">Name of the test case header template.</param>
        /// <param name="linkMapping">The link mapping.</param>
        /// <param name="progressService">The progress service.</param>
        /// <returns>
        /// Value indicating if the process has been canceled.
        /// </returns>
        private bool WriteWorkItemsToDocument(IWorkItemSyncAdapter sourceTfs,
            IWorkItemSyncAdapter destinationWord,
            ITestReportHelper testReportHelper,
            IDictionary<int, ITfsTestCaseDetail> testCases,
            Dictionary<IWorkItem, IWorkItem> mappings,
            SynchronizedCollection<IWorkItem> finished,
            string testCaseElementTemplateName,
            string testCaseHeaderTemplateName,
            Dictionary<IWorkItem, IWorkItem> linkMapping,
            IProgressService progressService)
        {
            var text = mappings.Values.Contains(null) ? Resources.GetWorkItems_ImportingOf : Resources.LO_Publish_05_SyncToWordOf;

            foreach (var pair in mappings)
            {
                // wait until the current work item is refreshed by a refresh thread.
                while (!finished.Contains(pair.Key))
                {
                    Thread.Sleep(100);
                }

                SyncServiceTrace.I(Resources.SynchronizationOfWorkItemToWord, pair.Key.Id);

                if (testCases.ContainsKey(pair.Key.Id))
                {
                    SyncServiceTrace.D(Resources.CreateWorkItem, pair.Key.Id);
                    testReportHelper.InsertHeaderTemplate(testCaseHeaderTemplateName);
                    testReportHelper.InsertTestCase(testCaseElementTemplateName, testCases[pair.Key.Id]);
                }
                else
                {
                    var destination = pair.Value;
                    if (destination == null)
                    {
                        SyncServiceTrace.D(Resources.CreateWorkItem, pair.Key.Id);
                        try
                        {
                            destination = destinationWord.CreateNewWorkItem(pair.Key.Configuration);
                        }
                        catch (ConfigurationException ce)
                        {
                            PublishException(ce, null, Resources.WorkItemSyncService_FailedToCreateNewWorkItem, pair.Key.Configuration.WorkItemType);
                        }
                    }
                    if (destination != null)
                    {
                        linkMapping.Add(destination, pair.Key);
                        SyncWorkItem(pair.Key, destination, sourceTfs, Direction.TfsToOther);
                    }
                }

                progressService.DoTick(string.Format(CultureInfo.CurrentCulture, text, progressService.ActualProgressGroupActualTick, progressService.ActualProgressGroupCountOfTicks));
                if (progressService.ProgressCanceled)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Writes the work items to the document.
        /// </summary>
        /// <param name="sourceTfs">The source TFS.</param>
        /// <param name="destinationWord">The destination word.</param>
        /// <param name="testReportHelper">The test report helper.</param>
        /// <param name="sharedSteps">The test cases.</param>
        /// <param name="mappings">The mappings.</param>
        /// <param name="finished">The finished.</param>
        /// <param name="testCaseElementTemplateName">Name of the test case element template.</param>
        /// <param name="testCaseHeaderTemplateName">Name of the test case header template.</param>
        /// <param name="linkMapping">The link mapping.</param>
        /// <param name="progressService">The progress service.</param>
        /// <returns>
        /// Value indicating if the process has been canceled.
        /// </returns>
        private bool WriteWorkItemsToDocument(IWorkItemSyncAdapter sourceTfs,
            IWorkItemSyncAdapter destinationWord,
            ITestReportHelper testReportHelper,
            IDictionary<int, ITfsSharedStepDetail> sharedSteps,
            Dictionary<IWorkItem, IWorkItem> mappings,
            SynchronizedCollection<IWorkItem> finished,
            string testCaseElementTemplateName,
            string testCaseHeaderTemplateName,
            Dictionary<IWorkItem, IWorkItem> linkMapping,
            IProgressService progressService)
        {
            var text = mappings.Values.Contains(null) ? Resources.GetWorkItems_ImportingOf : Resources.LO_Publish_05_SyncToWordOf;

            foreach (var pair in mappings)
            {
                // wait until the current work item is refreshed by a refresh thread.
                while (!finished.Contains(pair.Key))
                {
                    Thread.Sleep(100);
                }

                SyncServiceTrace.I(Resources.SynchronizationOfWorkItemToWord, pair.Key.Id);

                if (sharedSteps.ContainsKey(pair.Key.Id))
                {
                    SyncServiceTrace.D(Resources.CreateWorkItem, pair.Key.Id);
                    testReportHelper.InsertHeaderTemplate(testCaseHeaderTemplateName);
                    testReportHelper.InsertSharedStep(testCaseElementTemplateName, sharedSteps[pair.Key.Id]);
                }
                else
                {
                    var destination = pair.Value;
                    if (destination == null)
                    {
                        SyncServiceTrace.D(Resources.CreateWorkItem, pair.Key.Id);
                        try
                        {
                            destination = destinationWord.CreateNewWorkItem(pair.Key.Configuration);
                        }
                        catch (ConfigurationException ce)
                        {
                            PublishException(ce, null, Resources.WorkItemSyncService_FailedToCreateNewWorkItem, pair.Key.Configuration.WorkItemType);
                        }
                    }
                    if (destination != null)
                    {
                        linkMapping.Add(destination, pair.Key);
                        SyncWorkItem(pair.Key, destination, sourceTfs, Direction.TfsToOther);
                    }
                }

                progressService.DoTick(string.Format(CultureInfo.CurrentCulture, text, progressService.ActualProgressGroupActualTick, progressService.ActualProgressGroupCountOfTicks));
                if (progressService.ProgressCanceled)
                    return true;
            }
            return false;
        }

        private static void RefreshWorkItemsInParallel(int threadCount, Dictionary<IWorkItem, IWorkItem> mappings, SynchronizedCollection<IWorkItem> finished)
        {
            for (var i = 0; i < threadCount; i++)
            {
                // create new variables for thread closure. (You cannot remove threadIndex = i and use i in the closure until c# 5.0)
                var threadIndex = i;
                new Thread(() =>
                           {
                               foreach (var workItem in mappings.Keys.Where((x, index) => index % threadCount == threadIndex))
                               {
                                   workItem.Refresh();
                                   finished.Add(workItem);
                               }
                           }).Start();
            }
        }

        /// <summary>
        /// Synchronizes work items from source adapter to destination adapter. The mapping determines what items from the source adapter are mapped to what items in the destination adapter.
        /// </summary>
        /// <param name="mappings">Dictionary that maps work items in source adapter (keys) to work items in destination adapter (values). If the value is null, it is looked up from or created in the destination adapter</param>
        /// <param name="destinationAdapter">The adapter in which the work items are updated.</param>
        /// <param name="sourceAdapter">The adapter which is the source of the the mapping keys. This is needed to look up additional information about linked work items.</param>
        private void ReverseSync(Dictionary<IWorkItem, IWorkItem> mappings, IWorkItemSyncAdapter sourceAdapter, IWorkItemSyncAdapter destinationAdapter)
        {
            string text = mappings.Values.Contains(null) ? Resources.GetWorkItems_ImportingOf : Resources.LO_Publish_05_SyncToWordOf;
            var linkMapping = new Dictionary<IWorkItem, IWorkItem>();

            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.EnterProgressGroup(2);
            progressService.EnterProgressGroup(mappings.Count);

            // Create threads to refresh work items in parallel
            int threadCount = 4;
            var finished = new SynchronizedCollection<IWorkItem>();

            for (int i = 0; i < threadCount; i++)
            {
                // create new variables for thread closure. (You cannot remove threadIndex = i and use i in the closure until c# 5.0)
                int threadIndex = i;
                new Thread(() =>
                    {
                        foreach (var workItem in mappings.Keys.Where((x, index) => index % threadCount == threadIndex))
                        {
                            workItem.Refresh();
                            finished.Add(workItem);
                        }
                    }).Start();
            }

            foreach (var pair in mappings)
            {
                // wait until the current work item is refreshed by a refresh thread.
                while (!finished.Contains(pair.Key)) Thread.Sleep(100);

                SyncServiceTrace.I(Resources.SynchronizationOfWorkItemToWord, pair.Key.Id);
                Console.WriteLine(Resources.SynchronizationOfWorkItemToWord, pair.Key.Id);

                var destination = pair.Value;
                if (destination == null)
                {
                    destination = destinationAdapter.WorkItems.Find(pair.Key.Id);
                    if (destination == null)
                    {
                        SyncServiceTrace.D(Resources.CreateWorkItem, pair.Key.Id);
                        try
                        {
                            destinationAdapter.ProcessOperations(pair.Key.Configuration.PreOperations);
                            destination = destinationAdapter.CreateNewWorkItem(pair.Key.Configuration);
                            destinationAdapter.ProcessOperations(pair.Key.Configuration.PostOperations);
                        }
                        catch (Exception ce)
                        {
                            PublishException(ce, null, Resources.WorkItemSyncService_FailedToCreateNewWorkItem, pair.Key.Configuration.WorkItemType);
                        }
                    }
                }
                if (destination != null)
                {
                    linkMapping.Add(destination, pair.Key);
                    SyncWorkItem(pair.Key, destination, sourceAdapter, Direction.TfsToOther);
                }

                progressService.DoTick(string.Format(CultureInfo.CurrentCulture, text, progressService.ActualProgressGroupActualTick, progressService.ActualProgressGroupCountOfTicks));
                if (progressService.ProgressCanceled)
                    return;
            }
            progressService.LeaveProgressGroup();
            progressService.DoTick();
            LinkWorkItems(linkMapping, sourceAdapter, Direction.TfsToOther);
            progressService.LeaveProgressGroup();
        }

        /// <summary>
        /// Overwrite the flags set by the pre-check of the WorkItems in the Wordribbon.
        /// Loops through all items and updates the flags
        /// This is only done for WordTableFieldItems
        /// </summary>
        /// <param name="workItemsToPublish">The work items that are about to get published</param>
        /// <param name="workItems">The pre-checked list of workitems, the right flags exist only here</param>
        /// <returns>The updated list</returns>
        private IEnumerable<IWorkItem> OverwriteParseHTMLAsPlainTextFlag(IEnumerable<IWorkItem> workItemsToPublish, IEnumerable<IWorkItem> workItems)
        {
            //Users Decision is "true", set the flag for each field to trigger an overwrite
            //Loop through all WorkItems that should be published
            //Look if it is in the dictionary
            var selectedWorkItems = workItemsToPublish.Where(p => workItems.Any(p1 => p.IsSameWorkItem(p1)));

            foreach (var item in selectedWorkItems)
            {
                //Adjust all fields
                foreach (IField field in item.Fields)
                {
                    if (field is WordTableField && field.GetType() == typeof(WordTableField) && item.Fields[field.ReferenceName] != null)
                    {
                        ((WordTableField)(item.Fields[field.ReferenceName])).ParseHtmlAsPlaintext = ((WordTableField)field).ParseHtmlAsPlaintext;

                    }
                }
            }

            return workItemsToPublish;
        }

        /// <summary>
        /// Synchronizes Word requirements to TFS work items.
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private Dictionary<IWorkItem, IWorkItem> ForwardSync(IWorkItemSyncAdapter tfs, IWorkItemSyncAdapter word, IEnumerable<IWorkItem> workItems, bool forceOverwrite)
        {
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            var mappedItems = new Dictionary<IWorkItem, IWorkItem>();
            int stackRank = 1;

            progressService.EnterProgressGroup(workItems.Count());

            foreach (IWorkItem workItem in workItems)
            {
                SyncServiceTrace.I(Resources.SynchronizationOfWorkItemToTFS, workItem.Id);
                try
                {
                    IWorkItem destinationWorkItem = tfs.WorkItems.Find(workItem.Id);

                    // create new work item if there is no mapped item or the item is new
                    if (destinationWorkItem == null || workItem.IsNew)
                    {
                        destinationWorkItem = tfs.CreateNewWorkItem(workItem.Configuration);
                        SyncServiceTrace.D(Resources.CreateWorkItemInTFS);
                    }
                    else
                    {
                        // Do configured state transitions if neccessary and allowed
                        // Ignore the field "HierarchyLevel, as it is only used internal and should not synched to the TFS
                        var workItemToUse = workItem;
                        var config = _configuration.GetWorkItemConfigurationExtended(workItem.WorkItemType, fieldName =>
                        {
                            if (string.IsNullOrEmpty(fieldName))
                                return null;
                            if (workItemToUse.Configuration.WorkItemSubtypeField.Equals("HierarchyLevel"))
                                return workItemToUse.Configuration.WorkItemSubtypeValue;
                            var fieldToEvaluate = workItemToUse.Fields[fieldName];
                            return fieldToEvaluate == null ? null : fieldToEvaluate.Value;
                        });
                        foreach (var field in config.FieldConfigurations)
                        {
                            try
                            {
                                var name = field.ReferenceFieldName;
                                if (workItem.Fields.Contains(name) &&
                                    destinationWorkItem.Fields.Contains(name) &&
                                    !destinationWorkItem.Fields[name].CompareValue(workItem.Fields[name].Value, _configuration.IgnoreFormatting))
                                {
                                    // changed work item: state transition
                                    config.DoTransition(destinationWorkItem, workItem);
                                    break;
                                }
                            }
                            catch (KeyNotFoundException)
                            {
                                PublishWorkItemError(workItem, workItem.Fields[field.ReferenceFieldName], Resources.TextFormating_FieldDetail, field.ReferenceFieldName);
                                return mappedItems;
                            }

                        }

                        // Check for revision conflicts
                        if (!forceOverwrite)
                        {
                            //var conflicts = GetConflictingFields(workItem, destinationWorkItem);
                            var changedWordFields = SynchronizationInformationHelper.GetChangedWordFields(workItem, destinationWorkItem, _configuration.IgnoreFormatting);
                            var changedTfsFields = SynchronizationInformationHelper.GetChangedTfsFields(workItem, destinationWorkItem, _configuration.IgnoreFormatting);
                            var publishableWordFields = SynchronizationInformationHelper.GetPublishableWordFields(changedWordFields, destinationWorkItem);

                            var conflicts = SynchronizationInformationHelper.GetDivergedFields(publishableWordFields, changedTfsFields);
                            if (conflicts.Count > 0)
                            {
                                var info = PublishWorkItemError(workItem, null, Resources.TFSError_ConflictingChanges, string.Join("\r", conflicts.Select(x => x.ReferenceName).ToArray()));
                                info.Destination = destinationWorkItem;
                                info.Type = UserInformationType.Conflicting;

                                // continue with next work item, dont sync this one
                                continue;
                            }
                        }

                    }

                    // Increase stack rank. Do this only if multiple workitems are published.
                    if (_configuration.UseStackRank)
                    {
                        if (destinationWorkItem.Fields.Contains("Microsoft.VSTS.Common.StackRank"))
                        {
                            if (workItems.Count() > 1)
                            {
                                var rightValue = stackRank.ToString(CultureInfo.InvariantCulture);
                                destinationWorkItem.Fields["Microsoft.VSTS.Common.StackRank"].Value = rightValue;
                                stackRank++;
                            }

                        }
                    }

                    var success = SyncWorkItem(workItem, destinationWorkItem, word, Direction.OtherToTfs);

                    SyncServiceTrace.I(success ? Resources.LogService_Sync_SynchronizationSuccessful : Resources.LogService_Sync_SynchronizationFailed, workItem.Id);
                    if (success)
                    {
                        // TODO: Commented to avoid zombies in database
                        /*
                        if (!string.IsNullOrEmpty(word.Name))
                        {
                            if (word.Name.StartsWith(@"\\", StringComparison.OrdinalIgnoreCase) ||
                                word.Name.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                            {
                                destinationWorkItem.AddHyperlink(word.Name);
                            }
                        }
                         */
                        mappedItems.Add(destinationWorkItem, workItem);
                    }
                    else
                    {
                        workItem.HasWordValidationError = true;
                        destinationWorkItem.HasWordValidationError = true;
                    }

                    progressService.DoTick(string.Format(CultureInfo.CurrentCulture, Resources.LO_Publish_03_SyncToTFSOf, progressService.ActualProgressGroupActualTick, progressService.ActualProgressGroupCountOfTicks));
                    if (progressService.ProgressCanceled)
                        break;
                }
                catch (Exception e)
                {
                    PublishException(e, workItem, e.Message);
                }
            }

            progressService.LeaveProgressGroup();
            return mappedItems;
        }

        /// <summary>
        /// Synchronizes a source work item with a destination work item.
        /// </summary>
        /// <param name="source">Source work item.</param>
        /// <param name="destination">Destination work item.</param>
        /// <param name="sourceAdapter">The sourceAdapter is used when looking up hyperlinks to work items (TFS)</param>
        /// <param name="direction">Direction of operation.</param>
        /// <returns>Returns true on success or false on failure.</returns>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private bool SyncWorkItem(IWorkItem source, IWorkItem destination, IWorkItemSyncAdapter sourceAdapter, Direction direction)
        {
            var editedFields = new List<IField>();

            SyncServiceTrace.I(Resources.SynchronizationOfWorkItem, destination.Id, direction);

            int destinationRev = destination.Revision;

            var oleMarkerFieldValues = new Dictionary<string, HashSet<string>>();
            // Synchronize the FieldItems Word<->TFS
            foreach (var field in source.Fields)
            {
                try
                {
                    if (!(field is StaticValueField))
                    {
                        // Skip ConfigurationExcpetion for StackRank if not in destination, somhow this field is in source even if not mapped.
                        if (field.ReferenceName == FieldReferenceNames.StackRank && !destination.Fields.Any(x => x.ReferenceName == FieldReferenceNames.StackRank))
                            continue;

                        // There are unmapped fields in case of OLEmarker is used
                        // This check might avoid actually intended error logging
                        // a more precise condition would also check whether the field is came in the game because of the OLEMarker functionality
                        if (!destination.Fields.Contains(field.ReferenceName))
                            continue;

                        var destinationField = destination.Fields[field.ReferenceName];
                        var tfsFieldRev = source.GetFieldRevision(field.ReferenceName);
                        ;

                        string rightValue = field.Value;
                        string leftValue = string.Empty;

                        // Set up converters for AreaPath and IterationPath to cut the project name from their paths
                        if (field.ReferenceName == CoreFieldReferenceNames.AreaPath)
                            rightValue = _pathConverter.Convert(rightValue, direction);
                        if (field.ReferenceName == CoreFieldReferenceNames.IterationPath)
                            rightValue = _pathConverter.Convert(rightValue, direction);


                        var wordField = direction == Direction.OtherToTfs ? field : destinationField;

                        if (!wordField.Configuration.DateTimeFormat.IsNullOrEmpty() && !field.Value.IsNullOrEmpty() && !rightValue.IsNullOrEmpty())
                        {
                            var tempDate = field.OriginalValue as DateTime?;
                            if (tempDate == null)
                                tempDate = DateTime.Parse(rightValue);
                            rightValue = tempDate?.ToString(wordField.Configuration.DateTimeFormat);
                        }

                        SyncServiceTrace.D(Resources.SynchronizationOfField,
                            destination.Id,
                            field.ReferenceName,
                            wordField.Configuration.Direction,
                            wordField.Configuration.FieldValueType,
                            wordField.Configuration.HandleAsDocument,
                            wordField.Configuration.HandleAsDocumentMode);

                        // Syncing to TFS
                        if (Direction.OtherToTfs == direction)
                        {
                            SetOleMarkerFields(destination.Fields, wordField, oleMarkerFieldValues);

                            // only publish fields that are changed compared to the server version with the same revision.
                            var changedWordFields = SynchronizationInformationHelper.GetChangedWordFields(source, destination, _configuration.IgnoreFormatting);
                            var actualChangedFields = SynchronizationInformationHelper.GetPublishableWordFields(changedWordFields, destination);
                            if (!actualChangedFields.Any(x => x.ReferenceName.Equals(wordField.ReferenceName)))
                            {
                                continue;
                            }

                            // fields with publishOnly or OtherToTfs
                            if (wordField.Configuration.Direction == Direction.PublishOnly ||
                                wordField.Configuration.Direction == Direction.OtherToTfs ||
                                (wordField.Configuration.Direction == Direction.SetInNewTfsWorkItem && destination.IsNew))
                            {
                                //  // Images in word are displayed 20% smaller than they actually are.
                                //  // Thus we have to shrink the images to 80% in order to display them correctly in the TFS fields.
                                //  // On the other hand, when copied into word, image sizes are increased by 25%.
                                //  // Therefore they are again displayed correctly in word when we download them from the TFS.
                                SyncServiceTrace.D(Resources.CompareValues,
                                    destination.Id,
                                    destinationField.Value,
                                    rightValue,
                                    _configuration.IgnoreFormatting);

                                // check if a few requirements on word formatting are met to prevent
                                // publishing fields that are not exported correctly.
                                if (CheckWordFormattingRequirements(source, rightValue, field) == false)
                                {
                                    return false;
                                }

                                if (destinationField.CompareValue(rightValue, _configuration.IgnoreFormatting) == false)
                                {
                                    if (!(field.ReferenceName.Equals("Microsoft.VSTS.Common.StackRank") && _configuration.UseStackRank))
                                    {
                                        SyncServiceTrace.D(Resources.OverwritingWithNewValue, destination.Id);
                                        destinationField.Value = rightValue;
                                        editedFields.Add(field);
                                    }
                                }

                                destinationField.MicroDocument = field.MicroDocument;
                            }
                        }

                        // Syncing dropdownlist entries to word, even for publish only
                        if (Direction.TfsToOther == direction && wordField.Configuration.Direction == Direction.PublishOnly)
                        {
                            if (field.Configuration.FieldValueType == FieldValueType.DropDownList)
                            {
                                if (source.Fields[field.ReferenceName].AllowedValues.Contains(destinationField.Value))
                                {
                                    destination.Fields[field.ReferenceName].AllowedValues = source.Fields[field.ReferenceName].AllowedValues;
                                    destination.Fields[field.ReferenceName].Value = destinationField.Value;
                                }
                                else
                                {
                                    SyncServiceTrace.W(Resources.UpdateDropdownFailed);
                                }
                            }
                        }

                        // Syncing to word... always except case PublishOnly and GetOnly
                        if (Direction.TfsToOther == direction && wordField.Configuration.Direction != Direction.PublishOnly &&
                            !(wordField.Configuration.Direction == Direction.GetOnly && !destination.IsNew))
                        {
                            // this has to be done because the TFS converts some values of the string.
                            rightValue = rightValue.Replace('\r', '\n').Trim('\n');

                            try
                            {
                                leftValue = destinationField.Value;
                                SyncServiceTrace.D(Resources.OverwritingOldwithNewValue, source.Id, leftValue, rightValue);
                            }
                            catch (KeyNotFoundException)
                            {
                                SyncServiceTrace.D(Resources.OverwritingInvalidConverter, source.Id, rightValue);
                                // Invalid converter value
                            }

                            // check if values are actually different to make sure comments are not deleted unneccessarily
                            // broken formatting in word will not be updated with tfs formatting unless "ignore formatting" is deaktivated.
                            if (field.CompareValue(leftValue, _configuration.IgnoreFormatting) == false ||
                                field.Configuration.FieldValueType == FieldValueType.DropDownList ||
                                field.ReferenceName.Equals(CoreFieldReferenceNames.Id) ||
                                field.ReferenceName.Equals(CoreFieldReferenceNames.Rev))
                            {
                                // TODO compare micro documents
                                destination.Fields[field.ReferenceName].AllowedValues = source.Fields[field.ReferenceName].AllowedValues;
                                destination.Fields[field.ReferenceName].Value = rightValue;
                                editedFields.Add(field);

                                // Add bookmarks and hyperlinks for id and revision
                                if (field.ReferenceName.Equals(CoreFieldReferenceNames.Id))
                                {
                                    destinationField.Hyperlink = sourceAdapter.GetWorkItemEditorUrl(int.Parse(rightValue, CultureInfo.InvariantCulture)).ToString();
                                    destination.CreateBookmark();
                                }
                                else if (field.ReferenceName.Equals(CoreFieldReferenceNames.Rev))
                                {
                                    var tfsService = sourceAdapter as ITfsService;
                                    destination.Fields[field.ReferenceName].Hyperlink = tfsService.GetRevisionWebAccessUri(source, int.Parse(rightValue, CultureInfo.InvariantCulture)).ToString();
                                }
                            }

                            // handle special values
                            if (wordField.Configuration.HandleAsDocument)
                            {
                                // (-SER) why is this field only updated if the revision is higher? is it to prevent ole object to be overwritten
                                // with images in case HandleAsDocument is not set to true?
                                // (-BENNI F) the contents from the TFS are only overwritten if its revision is higher, except the case that both have the revision "0" (if they are both new). This was necessary to solve an issue related to ole types
                                if ((tfsFieldRev > destinationRev) || ((tfsFieldRev == 0) && (destinationRev == 0)))
                                {
                                    // set the temp field so when setting a non-existant special value, it instead uses the temp value
                                    destination.Fields[field.ReferenceName].Value = rightValue;
                                    destinationField.MicroDocument = field.MicroDocument;
                                }
                            }
                        }
                    }
                    else
                    {
                        SyncServiceTrace.D(Resources.SynchronizationStaticSkipped);
                    }
                }

                catch (FormatException)
                {
                    // error in link expanding
                    var wordItem = direction == Direction.OtherToTfs ? source : destination;
                    PublishWorkItemError(wordItem, wordItem.Fields[field.ReferenceName], Resources.TextFormating_FieldDetail, field.ReferenceName);
                    return false;
                }
                catch (KeyNotFoundException e)
                {
                    // converter value not found..maybe more, hard to tell
                    var wordItem = direction == Direction.OtherToTfs ? source : destination;
                    PublishWorkItemError(wordItem, wordItem.Fields[field.ReferenceName], e.Message, field.ReferenceName);
                    return false;
                }
                catch (ConfigurationException ex)
                {
                    // Reference to non existing field. This happens because a typo in a config file OR because some configured fields are not actually mapped - like stackRank or maybe state fields
                    var wordItem = direction == Direction.OtherToTfs ? source : destination;
                    PublishWorkItemError(wordItem,
                                         wordItem.Fields.Contains(field.ReferenceName)
                                             ? wordItem.Fields[field.ReferenceName]
                                             : wordItem.Fields[CoreFieldReferenceNames.Id],
                                         ex.Message);
                }
                catch (Exception ex)
                {
                    var wordItem = direction == Direction.OtherToTfs ? source : destination;
                    PublishWorkItemError(wordItem, wordItem.Fields[field.ReferenceName], ex.Message);
                    return false;
                }

            }

            // a value change may have caused fields to become read only (not editable)
            // Because values to these fields may have already been assigne before the field becomes readonly,
            // check again to make sure no such field was edited
            foreach (var field in editedFields)
            {
                if (destination.Fields[field.ReferenceName].IsEditable == false)
                {
                    if (destination.IsDirty)
                    {
                        destination.Refresh();
                    }

                    PublishWorkItemError(source, source.Fields[field.ReferenceName], Resources.FieldStatusInvalidNotOldValue2, field.ReferenceName);
                    return false;
                }
            }

            // Add the values for the static text fields
            foreach (var field in destination.Fields)
            {
                var valueField = field as StaticValueField;
                if (valueField != null)
                {

                    // Get all variables and search for the value of the variable in the field
                    var allVariables = _configuration.Variables;
                    var currentVariable = allVariables.FirstOrDefault(x => x.Name == valueField.VariableName);

                    // Nothing to do here, as a static Value field
                    if (currentVariable != null)
                    {
                        field.Value = currentVariable.Value;
                    }
                }

                if (field.Configuration.FieldValueType == FieldValueType.BasedOnVariable)
                {
                    // Get all variables and search for the value of the variable in the field
                    var allVariables = _configuration.Variables;
                    var currentVariable = allVariables.FirstOrDefault(x => x.Name == field.Configuration.VariableName);

                    // Nothing to do here, as a static Value field
                    if (currentVariable != null)
                    {
                        field.Value = currentVariable.Value;
                    }
                }
                else if (field.Configuration.FieldValueType == FieldValueType.BasedOnSystemVariable)
                {
                    string value;
                    if (!_systemVariableService.TryGetValueByName(field.Configuration.VariableName, out value))
                    {
                        PublishWorkItemError(source, source.Fields[field.ReferenceName], Resources.PublishUnknownSystemVariableName, field.Configuration.VariableName, field.Configuration.ReferenceFieldName);
                        return false;
                    }

                    field.Value = value;
                }
            }

            return true;
        }

        private void SetOleMarkerFields(IFieldCollection fields, IField wordField, Dictionary<string, HashSet<string>> oleFieldValues)
        {
            if (!String.IsNullOrEmpty(wordField.Configuration.OLEMarkerField))
            {
                var currentOLEMarkerField = wordField.Configuration.OLEMarkerField;
                currentOLEMarkerField.Trim();
                if (wordField.ContainsOleObject)
                {
                    var tempOLEMarkerValue = String.IsNullOrEmpty(wordField.Configuration.OLEMarkerValue) ? fields[wordField.ReferenceName].Name : wordField.Configuration.OLEMarkerValue.Trim();

                    if (!oleFieldValues.ContainsKey(currentOLEMarkerField))
                    {
                        oleFieldValues.Add(currentOLEMarkerField, new HashSet<string>());
                    }

                    var oleValues = oleFieldValues[currentOLEMarkerField];
                    oleValues.Add(tempOLEMarkerValue);
                    var valueForCustomField = string.Join(", ", oleValues);
                    fields[currentOLEMarkerField].Value = valueForCustomField;
                }
                else
                {
                    if (!oleFieldValues.ContainsKey(currentOLEMarkerField))
                        fields[currentOLEMarkerField].Value = string.Empty;
                }
            }
        }

        private static bool CheckWordFormattingRequirements(IWorkItem wordItem, string rightValue, IField field)
        {
            var wordField = field as WordTableField;
            if (wordField == null)
            {
                // TODO: Extend this check to split fields for headers
                return true;
            }

            // Prevent accidentially publishing a field with inline shapes that word fails to export.
            if (string.IsNullOrEmpty(rightValue) && wordField.ContainsShapes && wordField.Configuration.ShapeOnlyWorkaroundMode == ShapeOnlyWorkaroundMode.ShowAsError)
            {
                PublishWorkItemError(wordItem, wordItem.Fields[field.ReferenceName], Resources.PublishErrorInlineShapes, field.ReferenceName);
                return false;
            }

            // Warn when accidentially publishing a field with inline shapes in a text only field
            if (wordField.Configuration.FieldValueType != FieldValueType.HTML &&
                wordField.Configuration.HandleAsDocument == false && wordField.ContainsShapes)
            {
                PublishWorkItemError(wordItem, wordItem.Fields[wordField.ReferenceName], Resources.PublishWarningPlaintextInlineShapes, wordField.ReferenceName);
            }

            // Warn when publishing downscaled images as they have much lower quality.
            if (wordField.Configuration.FieldValueType == FieldValueType.HTML &&
                wordField.Configuration.HandleAsDocument == false && wordField.IsContainsScaledPictures)
            {
                var item = PublishWorkItemError(wordItem, wordItem.Fields[wordField.ReferenceName], Resources.PublishWarningDownscaledImages, wordField.ReferenceName);
                item.Type = UserInformationType.Warning;
            }

            // Warn when publishing a field with unclosed lists. this affects all html fields
            // except those that are always attached to the work item as micro documents
            if (wordField.Configuration.FieldValueType == FieldValueType.HTML &&
                (wordField.Configuration.HandleAsDocument == false || field.Configuration.HandleAsDocumentMode == HandleAsDocumentType.OleOnDemand) &&
                wordField.ContainsUnclosedLists)
            {
                PublishWorkItemError(wordItem, wordItem.Fields[field.ReferenceName], Resources.PublishWarningUnclosedLists, field.ReferenceName);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Syncs the links for all work items from source adapter to target adapter
        /// </summary>
        private void LinkWorkItems(Dictionary<IWorkItem, IWorkItem> mapping, IWorkItemSyncAdapter sourceAdapter, Direction direction)
        {
            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.EnterProgressGroup(mapping.Count);

            foreach (var pair in mapping)
            {
                if (progressService.ProgressCanceled) return;
                progressService.DoTick(string.Format(CultureInfo.InvariantCulture, Resources.SynchronizingLinks, pair.Key.Id));
                var destination = pair.Key;
                var source = pair.Value;

                if (source.Links == null)
                {
                    continue;
                }

                try
                {
                    if (direction == Direction.OtherToTfs)
                    {
                        LinkTfsWorkItems(source, destination, sourceAdapter, mapping);
                    }
                    else if (direction == Direction.TfsToOther)
                    {
                        LinkWordWorkItems(source, destination, sourceAdapter);
                    }
                }
                catch (ConfigurationException ce)
                {
                    PublishException(ce, direction == Direction.OtherToTfs ? source : destination, ce.Message);
                }
            }

            progressService.LeaveProgressGroup();
        }

        private void LinkWordWorkItems(IWorkItem source, IWorkItem destination, IWorkItemSyncAdapter sourceAdapter)
        {
            SyncServiceTrace.D(Resources.CreatingLinksForWordWorkItem, destination.Id);
            foreach (var sourceLinkItem in source.Links)
            {
                // To ensure link formatting for directions that don't sync to word, in these cases we add the already existing links again.
                var linkItem = sourceLinkItem;
                if ((sourceLinkItem.Key.Direction == Direction.PublishOnly && destination.IsNew == false) ||
                    (sourceLinkItem.Key.Direction == Direction.GetOnly && destination.IsNew == false))
                {
                    linkItem = destination.Links.FirstOrDefault(x => x.Key.LinkValueType == linkItem.Key.LinkValueType);
                    if (linkItem.Key == null)
                    {
                        continue;
                    }

                    // Expand links for fields that should not be refreshed only if word and tfs links are the same
                    if (linkItem.Value.OrderBy(x => x).SequenceEqual(sourceLinkItem.Value.OrderBy(x => x)) == false)
                    {
                        continue;
                    }
                }

                SyncServiceTrace.D(Resources.DeletingLinksOfType, destination.Id, linkItem.Key.LinkValueType);
                destination.AddLinks(sourceAdapter, new int[] { }, linkItem.Key.LinkValueType, linkItem.Key.LinkedWorkItemTypes, true);

                // Check if all referenced work items are loaded with all the fields needed for link formatting
                var requiredFields = linkItem.Key.GetLinkFormatRequiredFields().ToList();

                var neededIds = (from id in linkItem.Value
                                 let sourceWorkItem = sourceAdapter.WorkItems.Find(id)
                                 where sourceWorkItem == null || requiredFields.Any(f => !sourceWorkItem.Fields.Contains(f))
                                 select id).ToList();

                // If there are unloaded work items or work items that lack fields needed for link resolution, reload them.
                var adapter = sourceAdapter;
                if (neededIds.Any())
                {
                    var tfs = sourceAdapter as ITfsService;
                    adapter = SyncServiceFactory.CreateTfs2008WorkItemSyncAdapter(tfs.ServerName, tfs.ProjectName, null, _configuration);
                    adapter.Open(linkItem, requiredFields.Distinct().ToList());
                }

                // Filter by LinkedWorkItemTypes, if defined in WordToTFS template
                var linkedWorkItemTypes = linkItem.Key.GetLinkedWorkItemTypes();
                if (linkedWorkItemTypes != null && linkedWorkItemTypes.Any())
                {
                    var tmpIDs = new List<int>();
                    foreach (var tmpId in linkItem.Value)
                    {
                        var wi = adapter.WorkItems.Find(tmpId);
                        if (wi != null && linkedWorkItemTypes.Contains(wi.WorkItemType))
                        {
                            tmpIDs.Add(tmpId);
                        }
                    }

                    // todo: das neue Erstellen des KeyValuePairs kann noch ausgetauscht werden, indem direkt aus dem Int-Array in linkItem.Value die nicht gewünschten IDs entfernt werden
                    linkItem = new KeyValuePair<IConfigurationLinkItem, int[]>(linkItem.Key, tmpIDs.ToArray());
                }

                SyncServiceTrace.D(Resources.CreatingLinksOfType, destination.Id, linkItem.Key.LinkValueType, string.Join(",", linkItem.Value));
                destination.AddLinks(adapter, linkItem.Value, linkItem.Key.LinkValueType, linkItem.Key.LinkedWorkItemTypes, true);
            }

            // delete all link types in word that are not present in tfs work item.
            foreach (var linkItem in destination.Links.Where(x => !source.Links.Any(sourceLink => sourceLink.Key.LinkValueType == x.Key.LinkValueType)))
            {
                SyncServiceTrace.D(Resources.DeletingLinksOfType, destination.Id, linkItem.Key.LinkValueType);
                destination.AddLinks(sourceAdapter, new int[] { }, linkItem.Key.LinkValueType, true);
            }
        }

        private static void LinkTfsWorkItems(IWorkItem source, IWorkItem destination, IWorkItemSyncAdapter wordAdapter, Dictionary<IWorkItem, IWorkItem> mappings)
        {
            SyncServiceTrace.D(Resources.CreatingLinksForItemsOnTFS, destination.Id);

            foreach (var config in source.Configuration.Links)
            {
                if ((config.Direction == Direction.SetInNewTfsWorkItem && destination.IsNew == false) ||
                    config.Direction == Direction.GetOnly ||
                    config.Direction == Direction.TfsToOther)
                {
                    continue;
                }

                // add automatic links
                if (string.IsNullOrEmpty(config.AutomaticLinkWorkItemType) == false)
                {
                    var showMissingTargetWarning = true;
                    int targetId = 0;

                    for (var i = 0; i < wordAdapter.WorkItems.Count; i++)
                    {
                        var workItem = wordAdapter.WorkItems.ElementAt(i);

                        // only check work items that appear before the one we add links to.
                        if (workItem.Equals(source))
                        {
                            break;
                        }

                        // skip work items of the wrong type
                        if (!workItem.WorkItemType.Equals(config.AutomaticLinkWorkItemType))
                        {
                            continue;
                        }

                        // show error if subtype field is defined but not found
                        if (!string.IsNullOrEmpty(config.AutomaticLinkWorkItemSubtypeField) && !workItem.Fields.Contains(config.AutomaticLinkWorkItemSubtypeField))
                        {
                            PublishWorkItemError(source, null, Resources.PublishWarningMissingSubtypeField, config.AutomaticLinkWorkItemType, config.AutomaticLinkWorkItemSubtypeField);
                            showMissingTargetWarning = false;
                            break;
                        }

                        if (string.IsNullOrEmpty(config.AutomaticLinkWorkItemSubtypeField) ||
                            workItem.Fields[config.AutomaticLinkWorkItemSubtypeField].Value.Equals(config.AutomaticLinkWorkItemSubtypeValue))
                        {
                            // find the target id. If target is in publish set, use TFS work item id.
                            // If target is not in publish set use Word id (must not be new work item)
                            showMissingTargetWarning = false;
                            if (workItem.IsNew)
                            {
                                if (mappings.Values.Contains(workItem))
                                {
                                    targetId = mappings.Keys.First(x => mappings[x].Equals(workItem)).Id;
                                }
                                else
                                {
                                    PublishWorkItemError(source, null, Resources.PublishWarningNewAutoLinkTarget, workItem.Title);
                                    break;
                                }
                            }
                            else
                            {
                                targetId = workItem.Id;
                            }
                        }
                    }

                    // show warning if autolink is defined but no work item matching the criteria is found
                    if (targetId != 0)
                    {
                        SyncServiceTrace.D("WI{0} - Automatically creating links of type {1}. New ID={2} (Overwrite = {3})", destination.Id, config.LinkValueType, targetId, config.Overwrite);
                        destination.AddLinks(null, new[] { targetId }, config.LinkValueType, config.Overwrite);
                    }
                    else if (showMissingTargetWarning && config.AutomaticLinkSuppressWarnings == false)
                    {
                        var info = new UserInformation(source)
                        {
                            Explanation = string.Format(CultureInfo.CurrentCulture, Resources.PublishWarningNoTarget),
                            Type = UserInformationType.Warning,
                            NavigateToSourceAction = null,
                        };
                        Publish(info);
                    }
                }
                else
                {
                    var linkItem = source.Links.Where(x => x.Key.LinkValueType == config.LinkValueType).ToList();
                    if (linkItem.Any())
                    {
                        SyncServiceTrace.D("WI{0} - Creating links of type {1}. New ID's={2} (Overwrite = {3})", destination.Id, config.LinkValueType, string.Join(",", linkItem.First().Value), config.Overwrite);
                        destination.AddLinks(null, linkItem.First().Value, config.LinkValueType, config.LinkedWorkItemTypes, config.Overwrite);
                    }
                }
            }
        }

        /// <summary>
        /// Validates all work items in the destination adapter
        /// </summary>
        /// <returns>Returns whether all items are validated successful or not</returns>
        private static bool SyncValidateWorkItems(IWorkItemSyncAdapter destinationAdapter, IWorkItemSyncAdapter lookupAdapter)
        {
            var errorList = destinationAdapter.ValidateWorkItems();

            foreach (var error in errorList)
            {
                // look up the work item in the source adapter (word). Use the id to find the work item and the
                // title for new work items that caused an error while saving to tfs and have no id yet.
                var sourceWorkItem = error.WorkItem.Id == 0 ? lookupAdapter.WorkItems.FirstOrDefault(wi => wi.Title.Equals(error.WorkItem.Title)) : lookupAdapter.WorkItems.Find(error.WorkItem.Id);
                IField sourceField = null;

                if (sourceWorkItem != null && error.Field != null && sourceWorkItem.Fields.Contains(error.Field.ReferenceName))
                {
                    sourceField = sourceWorkItem.Fields[error.Field.ReferenceName];
                }

                PublishWorkItemError(sourceWorkItem ?? error.WorkItem, sourceField, error.Exception.Message);
            }

            return errorList.Count == 0;
        }

        /// <summary>
        /// The method saves all changes in the adapter. Removes work items with save errors from mappings
        /// </summary>
        /// <param name="adapter">Adapter that should be saved.</param>
        /// <param name="mappings">The mapping of work items. The keys have to be work items of the adapter to save and if a save fails for a work item, it is removed from the mapping.</param>
        /// <returns><c>True</c> on success. Otherwise <c>false</c>.</returns>
        private static bool SaveChanges(IWorkItemSyncAdapter adapter, Dictionary<IWorkItem, IWorkItem> mappings)
        {
            IList<ISaveError> errors;
            try
            {
                errors = adapter.Save();
            }
            catch (Exception ex)
            {
                PublishException(ex, null, Resources.TFSError_TFS_save_failed);
                return false;
            }

            foreach (var error in errors)
            {
                PublishWorkItemError(error.WorkItem, error.Field, error.Exception.Message);
                mappings.Remove(error.WorkItem);
            }

            return errors.Count == 0;
        }

        /// <summary>
        /// Publish the occurrence of an error that is attributed to a work item. This uses a standard formatting
        /// using the work item id and title
        /// </summary>
        /// <param name="source">The work item producing the error</param>
        /// <param name="field">The field where the error happened or null if not known</param>
        /// <param name="text">Helpful message to display</param>
        /// <param name="textFormatParameter">Parameters to insert into the message</param>
        /// <returns>The newly created publish information.</returns>
        private static IUserInformation PublishWorkItemError(IWorkItem source, IField field, string text, params object[] textFormatParameter)
        {
            var info = new UserInformation(source)
            {
                Explanation = string.Format(CultureInfo.CurrentCulture, text, textFormatParameter),
                Type = UserInformationType.Error,
                NavigateToSourceAction = field == null ? null : (Action)field.NavigateTo
            };
            SyncServiceTrace.W("{0}:{1}", info.Text, info.Explanation);
            Publish(info);

            return info;
        }

        /// <summary>
        /// Publish the occurrence of a generic error that is not attributed to a work item.
        /// </summary>
        /// <returns>The newly created publish information.</returns>
        private static void PublishException(Exception exception, IWorkItem source, string text, params object[] textFormatParameter)
        {
            var info = new UserInformation
            {
                Source = source,
                Text = string.Format(CultureInfo.CurrentCulture, text, textFormatParameter),
                Explanation = exception.Message,
                Type = UserInformationType.Error,
                NavigateToSourceAction = source == null ? null : (source.Fields.Contains(CoreFieldReferenceNames.Title) ? (Action)(() => source.Fields[CoreFieldReferenceNames.Title].NavigateTo()) : null)

            };
            SyncServiceTrace.LogException(exception);
            Publish(info);
        }

        /// <summary>
        /// Publish new information.
        /// </summary>
        private static void Publish(IUserInformation info)
        {
            var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
            if (infoStorage != null)
                infoStorage.AddItem(info);
        }

        #endregion
    }
}
