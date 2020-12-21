#region Usings
using AIT.TFS.SyncService.Adapter.TFS2012.BuildCenter;
using AIT.TFS.SyncService.Adapter.TFS2012.Helper;
using AIT.TFS.SyncService.Adapter.TFS2012.Properties;
using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
using AIT.TFS.SyncService.Adapter.TFS2012.WorkItemCollections;
using AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.BuildCenter;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Contracts.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Service.Configuration;
using AIT.TFS.SyncService.Service.InfoStorage;
using AIT.TFS.SyncService.Service.Utils;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Build.WebApi;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Common;
using Microsoft.TeamFoundation.ProcessConfiguration.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using FilterType = AIT.TFS.SyncService.Contracts.Enums.FilterType;
using WorkItemLinkFilter = AIT.TFS.SyncService.Service.Configuration.WorkItemLinkFilter;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012
{
    /// <summary>
    /// Adapter for team foundation server 2010.
    /// </summary>
    public class Tfs2012SyncAdapter : IWorkItemSyncAdapter, ITfsService, ITfsTestAdapter
    {
        #region Fields
        private readonly IConfiguration _configuration;
        private readonly IQueryConfiguration _queryConfiguration;
        private TfsTeamProjectCollection _tfs;
        private WorkItemStore _workItemStore;
        private List<ITFSWorkItemLinkType> _workItemLinkTypes;
        private IAreaIterationNode _areaPathNode;
        private IAreaIterationNode _iterationPathNode;
        private ITestManagementTeamProject _testManagement;
        private IBuildServer _buildServer;
        private IBuildDetail[] _builds;
        private Dictionary<int, IConfigurationItem> _items;
        private ICollection<WorkItem> _tfsItems;
        private IDictionary<string, WorkItemTypeAppearance> _colors;
        #endregion Fields

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="Tfs2012SyncAdapter"/> class.
        /// </summary>
        /// <param name="serverName">Name of the TFS server to work with.</param>
        /// <param name="projectName">Name of the TFS project to work with.</param>
        /// <param name="queryConfiguration">Configuration for query.</param>
        /// <param name="configuration">The mapping configuration to use when opening the adapter.</param>
        public Tfs2012SyncAdapter(string serverName, string projectName, IQueryConfiguration queryConfiguration, IConfiguration configuration)
        {
            Guard.ThrowOnArgumentNull(serverName, "serverName");
            Guard.ThrowOnArgumentNull(projectName, "projectName");

            ServerName = serverName;
            ProjectName = projectName;
            _queryConfiguration = queryConfiguration;
            _configuration = configuration;
            _tfsItems = new List<WorkItem>();
        }

        #endregion Constructor

        #region Private properties

        /// <summary>
        /// Gets the list of available work items from tfs.
        /// </summary>
        public ICollection<WorkItem> AvailableWorkItemsFromTFS
        {
            get
            {
                return _tfsItems;
            }
        }

        /// <summary>
        /// Gets the associated team project collection in this adapter.
        /// </summary>
        private TfsTeamProjectCollection TfsTeamProjectCollection
        {
            get
            {
                return _tfs ?? (_tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(ServerName, true, false));
            }
        }

        /// <summary>
        /// Gets the <see cref="WorkItemStore"/> from the actual team foundation server.
        /// </summary>
        private WorkItemStore WorkItemStore
        {
            get
            {
                if (null == _workItemStore)
                {
                    _workItemStore = (WorkItemStore)TfsTeamProjectCollection.GetService(typeof(WorkItemStore));
                    _workItemStore.RefreshCache();
                }

                return _workItemStore;
            }
        }


        /// <summary>
        /// Gets the associated team foundation server project in this adapter.
        /// </summary>
        private Project Project
        {
            get
            {
                if (null != WorkItemStore && WorkItemStore.Projects.Contains(ProjectName))
                {
                    return WorkItemStore.Projects[ProjectName];
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the <see cref="IBuildServer"/> to obtain build information.
        /// </summary>
        private IBuildServer BuildServer
        {
            get
            {
                if (_buildServer == null)
                {
                    _buildServer = TfsTeamProjectCollection.GetService<IBuildServer>();
                }
                return _buildServer;
            }
        }

        private IBuildDetail[] Builds
        {
            get
            {
                if (_builds == null)
                {
                    if (BuildServer == null)
                        return null;

                    IBuildDetailSpec buildSpec = BuildServer.CreateBuildDetailSpec(ProjectName);
                    buildSpec.InformationTypes = null;
                    _builds = BuildServer.QueryBuilds(buildSpec).Builds;
                }
                return _builds;
            }
        }
        #endregion Private properties

        #region Implementation of IWorkItemSyncAdapter interface

        /// <summary>
        /// A collection of work item colors.
        /// </summary>
        public IDictionary<string, WorkItemTypeAppearance> WorkItemColors
        {
            get
            {
                return _colors;
            }
        }

        /// <summary>
        /// A collection of all opened work items. The collection will be initialized when calling open
        /// </summary>
        public IWorkItemCollection WorkItems { get; private set; }

        /// <summary>
        /// Indicates whether this adapter has been opened. The list of work items will not be initialized unless
        /// an adapter is opened
        /// </summary>
        public bool IsOpen
        {
            get { return null != TfsTeamProjectCollection; }
        }

        /// <summary>
        /// Connects to the TFS adapter a queries the requested work items.
        /// </summary>
        /// <param name="ids">Ids of the work items to open</param>
        /// <returns>Returns true on success, false otherwise</returns>
        public bool Open(int[] ids)
        {
            return Open(ids, null);
        }

        /// <summary>
        /// Connects to the TFS adapter a queries the requested work items. Use this method if you want open each
        /// configured in the link format syntax)
        /// </summary>
        /// <param name="items">passed configurations</param>
        /// <returns>Returns true on success, false otherwise</returns>
        public bool OpenWithConfigurations(Dictionary<int, IConfigurationItem> items)
        {
            if (items == null)
            {
                throw new ArgumentNullException("items");
            }

            _items = items;
            return Open(items.Keys.ToArray(), null);
        }

        /// <summary>
        /// Connects to the TFS adapter a queries the requested work items. Use this method if you need to query more fields
        /// than only the fields specified in the work item mapping. (Link expanding uses this to query additional fields
        /// configured in the link format syntax)
        /// </summary>
        /// <param name="ids">Ids of the work items to open</param>
        /// <param name="fields">list of fields to query</param>
        /// <returns>Returns true on success, false otherwise</returns>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        public bool Open(int[] ids, IList<string> fields)
        {
            SyncServiceTrace.D("Connecting to the TFS adapter ...");
            if (IsOpen && WorkItems != null)
            {
                return true;
            }

            try
            {
                TfsTeamProjectCollection.EnsureAuthenticated();
                _colors = null;
                var _service = TfsTeamProjectCollection.GetService<ProjectProcessConfigurationService>();
                if (_service != null)
                {

                    _colors = new Dictionary<string, WorkItemTypeAppearance>();
                    try
                    {
                        try
                        {
                            var colorsFromWebApi = GetWitColorsFromWebApi();
                            _colors = colorsFromWebApi.ToDictionary(y => y.WorkItemTypeName);
                        }
                        catch
                        {                           
                            var colorsFromSoapApi = _service.GetProcessConfiguration(Project.Uri.AbsoluteUri).WorkItemColors.Select(x => new WorkItemTypeAppearance(x.WorkItemTypeName, x.PrimaryColor));
                            _colors = colorsFromSoapApi.ToDictionary(y => y.WorkItemTypeName);
                        }                      
                    }
                    catch
                    {
                        _colors = null;
                    }
                    finally
                    {
                        LoadWorkItems(ids, fields);
                    }
                }
            }
            catch (Exception ex)
            {
                SyncServiceTrace.LogException(ex);
                var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
                if (infoStorage != null)
                {
                    infoStorage.NotifyError(Resources.Error_ConnectTFS, ex);
                }
                else
                {
                    throw;
                }

                return false;
            }

            return true;
        }

        private IEnumerable<WorkItemTypeAppearance> GetWitColorsFromWebApi()
        {
            var vssCredentials = new Microsoft.VisualStudio.Services.Client.VssClientCredentials(true);
            var wiClient = new WorkItemTrackingHttpClient(_tfs.Uri, vssCredentials);
            var colorsfromWepApi= wiClient.GetWorkItemTypesAsync(Project.Name).Result.Select(wit => new WorkItemTypeAppearance(wit.Name, wit.Color));

            return colorsfromWepApi;
        }

        public bool Open(KeyValuePair<IConfigurationLinkItem, int[]> linkItem, IList<string> fields)
        {
            if (IsOpen && WorkItems != null)
            {
                return true;
            }

            try
            {
                TfsTeamProjectCollection.EnsureAuthenticated();
                LoadWorkItems(linkItem.Value, fields);
            }
            catch (Exception ex)
            {
                SyncServiceTrace.LogException(ex);
                var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
                if (infoStorage != null)
                {
                    infoStorage.NotifyError(Resources.Error_ConnectTFS, ex);
                }
                else
                {
                    throw;
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// Disconnects from the server and resets list of opened work items.
        /// </summary>
        public void Close()
        {
            if (!IsOpen)
            {
                WorkItems = null;
                if (null != TfsTeamProjectCollection)
                {
                    TfsTeamProjectCollection.Dispose();
                }
            }
        }

        /// <summary>
        /// Creates a new <see cref="IWorkItem"/> object in this adapter.
        /// </summary>
        /// <param name="configuration">Configuration of the work item to create</param>
        /// <returns>The new <see cref="IWorkItem"/> object or null if the adapter failed to create work item.</returns>
        /// <exception cref="ConfigurationException">Thrown when the adapter does not support the work item type</exception>
        public IWorkItem CreateNewWorkItem(IConfigurationItem configuration)
        {
            if (configuration == null) throw new ArgumentNullException("configuration");
            if (!Project.WorkItemTypes.Contains(configuration.WorkItemType))
            {
                throw new ConfigurationException(string.Format(CultureInfo.CurrentCulture, Resources.WorkItemIsNotAvailable, configuration.WorkItemTypeMapping));
            }

            var type = Project.WorkItemTypes[configuration.WorkItemType];
            var workItem = new TfsWorkItem(new WorkItem(type), this, null, configuration);
            WorkItems.Add(workItem);
            return workItem;
        }

        /// <summary>
        /// Saves all changes made to work items to its data source
        /// </summary>
        /// <returns>Returns a list of <see cref="ISaveError"/> if errors occurred. The list is empty if no error occurred</returns>
        public IList<ISaveError> Save()
        {
            // first save pass to upload changes and image attachments
            var result = InternalSave(WorkItems.Cast<TfsWorkItem>().ToList());
            if (result.Count > 0)
            {
                return result;
            }

            foreach (var item in WorkItems)
            {
                var workItem = (TfsWorkItem)item;
                workItem.SetHtmlValuesWithUpdatedImageSources();
            }

            // save the updated values and html uris
            result = InternalSave(WorkItems.Cast<TfsWorkItem>().ToList());

            // delete temporary files after save
            foreach (var workItem1 in WorkItems)
            {
                var workItem = (TfsWorkItem)workItem1;
                workItem.DeleteTempFilesAfterSave();
            }

            return result;
        }

        /// <summary>
        /// Validates all work items of the adapter to meet designated rules
        /// </summary>
        /// <returns>Returns a list of <see cref="ISaveError"/> if errors occurred. The list is empty if no error occurred</returns>
        public IList<ISaveError> ValidateWorkItems()
        {
            var errorList = new List<ISaveError>();
            foreach (var item in WorkItems)
            {
                var workItem = (TfsWorkItem)item;
                var errors = workItem.WorkItem.Validate();
                foreach (var error in errors)
                {
                    var tfsErrorField = error as Microsoft.TeamFoundation.WorkItemTracking.Client.Field;
                    if (tfsErrorField != null)
                    {
                        var errorField = (from field in workItem.Fields
                                          where field.ReferenceName.Equals(tfsErrorField.ReferenceName, StringComparison.OrdinalIgnoreCase)
                                          select field).FirstOrDefault();
                        errorList.Add(new SaveError(workItem, new ValidateWorkItemException(ErrorHelper.GetFieldStatusMessage(tfsErrorField))) { Field = errorField });
                    }
                }
            }

            return errorList;
        }

        /// <summary>
        /// Gets the list of all configuration field items that are available
        /// on the server.
        /// </summary>
        public FieldDefinitionCollection FieldDefinitions
        {
            get
            {
                return WorkItemStore.FieldDefinitions;
            }
        }

        /// <summary>
        /// Gets a web access URL for the work item editor
        /// </summary>
        /// <param name="workItemId">Id of the work item for which to create a web access URL</param>
        public Uri GetWorkItemEditorUrl(int workItemId)
        {
            var tswaClientHyperlinkService = TfsTeamProjectCollection.GetService<TswaClientHyperlinkService>();
            return tswaClientHyperlinkService.GetWorkItemEditorUrl(workItemId);
        }

        /// <summary>
        /// Process operations. This behavior is not implemented for the TFS adapter
        /// </summary>
        /// <param name="operations"></param>
        public void ProcessOperations(IList<IConfigurationTestOperation> operations)
        {
            // not implemented
        }

        #endregion Implementation of IWorkItemSyncAdapter interface

        #region Implementation of ITFSService
        /// <summary>
        /// Get all queries for current project.
        /// </summary>
        /// <returns>Project query hierarchy or null when failing to load query hierarchy</returns>
        public QueryHierarchy GetWorkItemQueries()
        {
            if (Project != null)
            {
                try
                {
                    Project.QueryHierarchy.Refresh();
                    return Project.QueryHierarchy;
                }
                catch (Exception ex)
                {
                    SyncServiceTrace.LogException(ex);
                    var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
                    if (infoStorage != null)
                    {
                        infoStorage.NotifyError(Resources.Error_GetQueryHierarchy, ex);
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Get link types for current project.
        /// </summary>
        /// <returns></returns>
        public ICollection<ITFSWorkItemLinkType> GetLinkTypes()
        {
            if (_workItemLinkTypes == null && WorkItemStore != null)
            {
                _workItemLinkTypes = new List<ITFSWorkItemLinkType>();
                foreach (var workItemType in WorkItemStore.WorkItemLinkTypes)
                {
                    _workItemLinkTypes.Add(new TfsWorkItemLinkType(workItemType, Contracts.WorkItemObjects.LinkType.ForwardEnd));
                    if (workItemType.IsDirectional)
                    {
                        _workItemLinkTypes.Add(new TfsWorkItemLinkType(workItemType, Contracts.WorkItemObjects.LinkType.ReverseEnd));
                    }

                }
            }
            return _workItemLinkTypes;
        }

        /// <summary>
        /// Gets query for query item definition.
        /// </summary>
        /// <param name="queryItem">Query item definition.</param>
        /// <returns></returns>
        private Query GetQuery(QueryItem queryItem)
        {
            var definition = queryItem as QueryDefinition;
            if (definition != null)
            {
                var wiqlContext = new Dictionary<string, object>
                                      {
                                          { "project", ProjectName },
                                          { "me", WorkItemStore.TeamProjectCollection.AuthorizedIdentity.DisplayName },
                                          { "today", DateTime.Now.Date }
                                      };
                return new Query(WorkItemStore, definition.QueryText, wiqlContext);
            }
            return null;
        }

        /// <summary>
        /// Get root of area tree.
        /// </summary>
        public IAreaIterationNode GetAreaPathHierarchy()
        {
            if (_areaPathNode == null)
            {
                _areaPathNode = new AreaIterationNode(GetStructureName(StructureType.ProjectModelHierarchy), Project.Name, string.Empty, null);
                GetAreaIterationNodes(_areaPathNode, Project.AreaRootNodes, _areaPathNode);
            }

            return _areaPathNode;
        }

        /// <summary>
        /// Get root of iteration tree.
        /// </summary>
        public IAreaIterationNode GetIterationPathHierarchy()
        {
            if (_iterationPathNode == null)
            {
                _iterationPathNode = new AreaIterationNode(GetStructureName(StructureType.ProjectLifecycle), Project.Name, string.Empty, null);
                GetAreaIterationNodes(_iterationPathNode, Project.IterationRootNodes, _iterationPathNode);
            }

            return _iterationPathNode;
        }

        /// <summary>
        /// Get all work item types for current project.
        /// </summary>
        /// <returns></returns>
        public WorkItemTypeCollection GetAllWorkItemTypes()
        {
            return Project.WorkItemTypes;
        }

        /// <summary>
        /// Gets revision web access url template.
        /// </summary>
        public Uri GetRevisionWebAccessUri(IWorkItem workItem, int revision)
        {
            Guard.ThrowOnArgumentNull(workItem, "workItem");

            const string Template = "{0}/WorkItemTracking/workitem.aspx?artifactMoniker={1}&Rev={2}";
            var replacedTemplate = string.Format(CultureInfo.InvariantCulture, Template, TfsTeamProjectCollection.Uri.AbsoluteUri, workItem.Id, revision);

            return new Uri(replacedTemplate);
        }

        /// <summary>
        /// Gets the category collection of the current project.
        /// </summary>
        /// <value>
        /// The category collection.
        /// </value>
        public CategoryCollection CategoryCollection
        {
            get
            {
                return Project == null ? null : Project.Categories;
            }
        }

        /// <summary>
        /// Name of the team project the adapter is connected to
        /// </summary>
        public string ProjectName { get; private set; }

        /// <summary>
        /// Name / URL of the TFS server
        /// </summary>
        public string ServerName { get; private set; }

        #endregion Implementation of ITFSService

        #region Implementation of ITfsTestAdapter

        /// <summary>
        /// Gets the <see cref="ITestManagementTeamProject"/> to obtain test manager information.
        /// </summary>
        public ITestManagementTeamProject TestManagement
        {
            get
            {
                if (_testManagement == null)
                {
                    _testManagement = TfsTeamProjectCollection.GetService<ITestManagementService>().GetTeamProject(ProjectName);
                }
                return _testManagement;
            }
        }

        /// <summary>
        /// Gets the available test plans.
        /// </summary>
        public IList<ITfsTestPlan> AvailableTestPlans
        {
            get
            {
                if (TestManagement == null)
                {
                    return null;
                }
                // Select all test plans.
                var testPlans = from plan in TestManagement.TestPlans.Query("SELECT * FROM TestPlan") where plan != null select plan;
                var retValue = new List<ITfsTestPlan>();

                // Iterate selected test plans.
                foreach (var item in testPlans)
                {
                    // It is requred that only active elements are selected.
                    // ReSharper disable once CSharpWarnings::CS0618
                    if (item.State != TestPlanState.Active)
                    {
                        continue;
                    }

                    // Create instance of wrap class.
                    var testPlan = new TfsTestPlan(item);

                    // Add to the list.
                    retValue.Add(testPlan);
                }
                if (retValue.Count == 0)
                {
                    SyncServiceTrace.D(Resources.AvailableTestPlanNotExists);
                }
                return retValue;
            }
        }

        /// <summary>
        /// Gets the available test configurations.
        /// </summary>
        public IList<ITfsTestConfiguration> AvailableTestConfigurations
        {
            get
            {
                if (TestManagement == null)
                {
                    return null;
                }
                var listOfTestConfigurations = from config in TestManagement.TestConfigurations.Query("SELECT * FROM TestConfiguration")
                                               where config != null
                                               select config;
                var retValue = new List<ITfsTestConfiguration>();
                foreach (var element in listOfTestConfigurations)
                {
                    var testConfiguration = new TfsTestConfiguration(element);
                    retValue.Add(testConfiguration);
                }
                if (retValue.Count == 0)
                {
                    SyncServiceTrace.D(Resources.AvailableTestConfigurationNotExists);
                }
                return retValue;
            }
        }

        /// <summary>
        /// Gets the available server builds.
        /// </summary>
        public IList<ITfsServerBuild> GetAvailableServerBuilds()
        {
            var builds = GetAllBuilds();
            return ApplyBuildFilters(builds);
        }

        /// <summary>
        /// Gets a list of all builds.
        /// </summary>
        /// <returns>List of new and old build</returns>
        private List<ITfsServerBuild> GetAllBuilds()
        {
            var buildTaskList = new List<ITfsServerBuild>();

            if (BuildServer == null)
            {
                return null;
            }
            try
            {
                // get Xaml Builds via the "old" API model (SOAP based)
                var buildSpec = BuildServer.CreateBuildDetailSpec(ProjectName);
                // Set the InformationTypes to null - Query need no more 4 minutes, but 2 seconds.
                buildSpec.InformationTypes = null;
                var xamlBuilds = BuildServer.QueryBuilds(buildSpec).Builds;
                foreach (var xamlBuild in xamlBuilds)
                {
                    buildTaskList.Add(new TfsServerBuild(xamlBuild));
                }

                // get Json Builds via the "new" API model (REST based)
                // -> since NuGet package Microsoft.TeamFoundationServer.ExtendedClient version 15.*, the 4.x API is used which does not support Xaml anymore
                var vssCredentials = new Microsoft.VisualStudio.Services.Client.VssClientCredentials(true);
                var buildClient = new BuildHttpClient(BuildServer.TeamProjectCollection.Uri, vssCredentials);
                var builds = buildClient.GetBuildsAsync(ProjectName).Result;
                foreach (var build in builds)
                {
                    buildTaskList.Add(new TfsServerBuild(build));
                }

                // if neither Xaml nor Json builds available, trace an approprate message
                if (buildTaskList.Count == 0)
                {
                    SyncServiceTrace.D(Resources.AvailableBuildNotExists);
                }
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                {
                    SyncServiceTrace.D(e.InnerException.Message);
                    var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
                    infoStorage.AddItem(new UserInformation { Text = string.Format(CultureInfo.InvariantCulture, e.Message), Explanation = e.InnerException == null ? string.Empty : e.InnerException.Message, Type = UserInformationType.Warning });
                }
                else
                {
                    throw;
                }
            }

            return buildTaskList;
        }

        public IList<ITfsServerBuild> ApplyBuildFilters(List<ITfsServerBuild> builds)
        {
            if (builds == null)
            {
                return null;
            }
            var retValue = new List<ITfsServerBuild>();

            var buildQualities = _configuration.ConfigurationTest.ConfigurationTestResult.BuildQualities;
            var buildFilters = _configuration.ConfigurationTest.ConfigurationTestResult.BuildFilters;
            IList<string> buildNames = null;
            string buildAge = null;
            IList<string> buildQualitiesFromFilter = null;
            IList<string> buildTags = null;

            if (buildFilters != null)
            {
                buildNames = _configuration.ConfigurationTest.ConfigurationTestResult.BuildFilters.BuildNames;
                buildAge = _configuration.ConfigurationTest.ConfigurationTestResult.BuildFilters.BuildAge;
                buildQualitiesFromFilter = _configuration.ConfigurationTest.ConfigurationTestResult.BuildFilters.BuildQualities;
                buildTags = _configuration.ConfigurationTest.ConfigurationTestResult.BuildFilters.BuildTags;
            }
            //conditions
            var buildFiltersExist = buildFilters != null;
            var buildNamesExist = buildFiltersExist && buildNames != null && buildNames.Count > 0;
            var buildAgeExist = buildFiltersExist && buildAge != null;
            var buildQualitesExist = (buildQualities != null && buildQualities.Count > 0) ||
                                     (buildQualitiesFromFilter != null && buildQualitiesFromFilter.Count > 0);
            var buildTagsExist = buildFiltersExist && buildTags != null && buildTags.Count > 0;

            //validation
            //case : buildAge exists and it is empty
            if (buildAgeExist && buildAge.Equals(string.Empty))
            {
                buildAgeExist = false;
            }

            //case : buildFilters exists and buildQualities filter is out of them
            if (buildFiltersExist && buildQualities != null && buildQualities.Count > 0)
            {
                throw new Exception(string.Format(Resources.Error_BuildQualityIsNotInsideBuildFilters));
            }

            double dBuildAge = 0;
            if (!string.IsNullOrWhiteSpace(buildAge) && !double.TryParse(buildAge, out dBuildAge))
            {
                throw new Exception(string.Format(Resources.Error_BuildAgeDouble));
            }

            foreach (var build in builds)
            {
                if (build == null)
                {
                    continue;
                }

                var isBuildQualityAcceptable = false;
                var isBuildNameAcceptable = false;
                var isBuildAgeAcceptable = false;
                var isBuildTagAcceptable = false;

                // for the old xaml builds we apply quality filters
                switch (build.DefinitionType)
                {
                    case DefinitionType.Xaml:
                        {
                            isBuildTagAcceptable = true;
                            // filter by build qualities
                            var qualities = buildFilters != null ? buildQualitiesFromFilter : buildQualities;

                            if (build.Quality != null && qualities != null)
                            {
                                foreach (var pattern in qualities)
                                {
                                    isBuildQualityAcceptable = string.IsNullOrEmpty(pattern) || (Regex.IsMatch(build.Quality, pattern));
                                    if (isBuildQualityAcceptable) break;
                                }
                            }
                            break;
                        }
                    case DefinitionType.Build:
                        {
                            isBuildQualityAcceptable = true;
                            foreach (var buildTag in build.Tags)
                            {
                                isBuildTagAcceptable = buildTags != null && buildTags.Any(x => x.Equals(buildTag));
                                if (isBuildTagAcceptable) break;
                            }
                            break;
                        }
                }

                // filter by build names
                isBuildNameAcceptable = buildNamesExist && buildNames.Contains(build.BuildNumber);

                // filter by build age
                isBuildAgeAcceptable = buildAgeExist && build.FinishTime.Value.AddDays(dBuildAge) >= DateTime.Today;

                // filter all together
                var isBuildValid = (isBuildQualityAcceptable && isBuildNameAcceptable && isBuildAgeAcceptable && isBuildTagAcceptable) ||
                                   //1 not exist, 3 exists
                                   (!buildAgeExist && isBuildQualityAcceptable && isBuildNameAcceptable && isBuildTagAcceptable) ||
                                   (!buildNamesExist && isBuildQualityAcceptable && isBuildAgeAcceptable && isBuildTagAcceptable) ||
                                   (!buildQualitesExist && isBuildNameAcceptable && isBuildAgeAcceptable && isBuildTagAcceptable) ||
                                   (!buildTagsExist && isBuildNameAcceptable && isBuildAgeAcceptable && isBuildQualityAcceptable) ||

                                   //2 not exist, 2 exists
                                   (isBuildQualityAcceptable && isBuildNameAcceptable && !buildAgeExist && !buildTagsExist) ||
                                   (isBuildQualityAcceptable && !buildNamesExist && isBuildAgeAcceptable && !buildTagsExist) ||
                                   (isBuildQualityAcceptable && !buildNamesExist && !buildAgeExist && isBuildTagAcceptable) ||
                                   (!buildQualitesExist && isBuildNameAcceptable && isBuildAgeAcceptable && !buildTagsExist) ||
                                   (!buildQualitesExist && isBuildNameAcceptable && !buildAgeExist && isBuildTagAcceptable) ||
                                   (!buildQualitesExist && !buildNamesExist && isBuildAgeAcceptable && isBuildTagAcceptable) ||

                                   //3 not exist, 1 exists
                                   (isBuildQualityAcceptable && !buildNamesExist && !buildAgeExist && !buildTagsExist) ||
                                   (!buildQualitesExist && isBuildNameAcceptable && !buildAgeExist && !buildTagsExist) ||
                                   (!buildQualitesExist && !buildNamesExist && isBuildAgeAcceptable && !buildTagsExist) ||
                                   (!buildQualitesExist && !buildNamesExist && !buildAgeExist && isBuildTagAcceptable);


                if (isBuildValid)
                {
                    retValue.Add(build);
                }
            }
            return retValue.OrderBy(o => o.BuildNumber).ToList();
        }

        /// <summary>
        /// The method returns all available test plans with flag 'Active' and test results for given server build.
        /// </summary>
        /// <param name="serverBuild">Server build to use as filter. If <c>null</c>, parameter is ignored.</param>
        /// <returns>Available test plans.</returns>
        public IList<ITfsTestPlan> GetTestPlans(ITfsServerBuild serverBuild)
        {
            // In some cases is TestCaseResult.BuildNumber 'null', but result is joined with build.
            // In this case is the connection over TestRun provided.
            var testPlans = AvailableTestPlans;
            if (serverBuild == null)
            {
                return testPlans;
            }

            // Create return list
            var retValue = new List<ITfsTestPlan>();

            // Get all test runs associated with this build
            var selectTestRuns = TestManagement.TestRuns.ByBuild(serverBuild.Uri);
            // If no test rund exists, exit method
            var allTestRuns = selectTestRuns as IList<ITestRun> ?? selectTestRuns.ToList();
            if (!allTestRuns.Any())
            {
                return retValue;
            }
            SyncServiceTrace.D(Resources.NumberOfAvailableTestPlans + testPlans.Count);
            SyncServiceTrace.D(Resources.NumberOfAvailableTestRuns + allTestRuns.Count);
            foreach (var tfsTestPlan in testPlans)
            {
                var testPlan = (TfsTestPlan)tfsTestPlan;
                var isEnableToAddTestPlan = false;

                // Get all test cases for test plan
                var allTestCases = testPlan.OriginalTestPlan.RootSuite.AllTestCases;
                // Iterate through test runs
                foreach (var testRun in allTestRuns)
                {
                    // Select test results for test run
                    var testResults = TestManagement.TestResults.Query($"SELECT * FROM TestResult WHERE TestRunId = {testRun.Id}");
                    // Check selected test results
                    if (allTestCases.Any(testCase => testResults.Any(testResult => testResult.TestCaseId == testCase.Id)))
                    {
                        isEnableToAddTestPlan = true;
                        break;
                    }
                }

                if (isEnableToAddTestPlan)
                {
                    retValue.Add(testPlan);
                }
            }

            return retValue;
        }

        /// <summary>
        /// Gets the available test configurations for given test plan.
        /// </summary>
        /// <param name="serverBuild">Server build to use as filter. If <c>null</c>, parameter is ignored.</param>
        /// <param name="testPlan">Test plan to get the available test configurations for</param>
        /// <returns>Available test configurations for given test plan.</returns>
        public IList<ITfsTestConfiguration> GetTestConfigurationsForTestPlanWithTestResults(ITfsServerBuild serverBuild, ITfsTestPlan testPlan)
        {
            if (testPlan == null)
            {
                throw new ArgumentNullException("testPlan");
            }
            var testConfigurations = AvailableTestConfigurations;
            var retValue = new List<ITfsTestConfiguration>();
            var tfstestPlan = _testManagement.TestPlans.Find(testPlan.Id);
            var testPoints = tfstestPlan.QueryTestPoints("SELECT * FROM TestPoint");

            foreach (var tp in testPoints)
            {
                if (tp.MostRecentResult != null && !retValue.Any(x => x.Id == tp.ConfigurationId))
                {
                    retValue.Add(testConfigurations.First(x => x.Id == tp.ConfigurationId));
                }
            }
            return retValue;
        }


        /// <summary>
        /// Gets the available test configurations for given test plan.
        /// </summary>
        /// <param name="testPlan">Test plan to get the available test configurations for</param>
        /// <returns>Available test configurations for given test plan.</returns>
        public IList<ITfsTestConfiguration> GetAllTestConfigurationsForTestPlan(ITfsTestPlan testPlan)
        {
            SyncServiceTrace.D(Resources.GetAllTestConfigurationsForTestPlan);
            if (testPlan == null)
            {
                SyncServiceTrace.D(Resources.TestPlanNotExists);
                throw new ArgumentNullException("testPlan");
            }
            var testConfigurations = AvailableTestConfigurations;
            var retValue = new List<ITfsTestConfiguration>();
            var tfstestPlan = _testManagement.TestPlans.Find(testPlan.Id);
            var testPoints = tfstestPlan.QueryTestPoints("SELECT * FROM TestPoint");

            foreach (var tp in testPoints)
            {
                if (!retValue.Any(x => x.Id == tp.ConfigurationId))
                {
                    retValue.Add(testConfigurations.First(x => x.Id == tp.ConfigurationId));
                }
            }
            return retValue;
        }

        /// <summary>
        /// Gets a signle <see cref="IWorkItem"/> using the work item ID and the current adapter.
        /// </summary>
        /// <param name="id">The work item id.</param>
        /// <returns></returns>
        public IWorkItem GetTfsWorkItemById(int id)
        {
            SyncServiceTrace.D(Resources.GetTfsWorkItemById);
            var realfsWi = WorkItemStore.GetWorkItem(id);
            var mappedConfiguration = _configuration.GetConfigurationItems().FirstOrDefault(x => x.WorkItemType == realfsWi.Type.Name);
            if (mappedConfiguration == null)
            {
                SyncServiceTrace.E(string.Format(Resources.NoMappingEntryFound, realfsWi.Type.Name, id));
                throw new ConfigurationException(string.Format(Resources.NoMappingEntryFound, realfsWi.Type.Name, id));
            }

            return new TfsWorkItem(realfsWi, this, null, mappedConfiguration);
        }

        /// <summary>
        /// Gets the available test configurations for given test plan. That are assigned to any TestCase in a TestPlan
        /// </summary>
        /// <param name="testPlan">Test plan to get the available test configurations for</param>
        /// <param name="latest">Determine if only the configigurations of the latest test results should be obtained</param>
        /// <param name="allConfigurations">Determine if all configuration </param>
        /// <param name="expandSharedSteps">Determines whether SharedSteps shall be shown expander or collapsed.</param>
        /// <returns>Available teGst configurations for given test plan.</returns>
        public IList<ITfsTestConfiguration> GetAssignedTestConfigurationsForTestPlan(ITfsTestPlan testPlan, bool latest, bool allConfigurations, bool expandSharedSteps)
        {
            SyncServiceTrace.D(Resources.GetAssignedTestConfigurationsForTestPlan);
            if (testPlan == null)
            {
                SyncServiceTrace.D(Resources.TestPlanNotExists);
                throw new ArgumentNullException("testPlan");
            }
            var testConfigurations = AvailableTestConfigurations;
            var retValue = new List<ITfsTestConfiguration>();
            foreach (var testConfiguration in testConfigurations)
            {
                if (TestConfigurationIsAssigned(testPlan, GetTestCases(testPlan, expandSharedSteps), testConfiguration, latest, allConfigurations, null))
                {
                    retValue.Add(testConfiguration);
                }
            }

            return retValue;
        }

        /// <summary>
        /// Gets the available test configurations for given test plan. That are assigned to any TestCase in a TestPlan
        /// </summary>
        /// <param name="testPlan">Test plan to get the available test configurations for</param>
        /// <param name="testCases"></param>
        /// <param name="mostRecentTestResults">Determine if only the test configurations of the latest test results should be obtained</param>
        /// <param name="allConfigurations">Determine if all configuration </param>
        /// <param name="tfsTestSuite"></param>
        /// <returns>Available test configurations for given test plan.</returns>
        public IEnumerable<ITfsTestConfigurationDetail> GetAssignedTestConfigurationsForTestCases(ITfsTestPlan testPlan, IList<ITfsTestCaseDetail> testCases, bool mostRecentTestResults, bool allConfigurations, ITfsTestSuite tfsTestSuite)
        {

            SyncServiceTrace.D(Resources.GetAssignedTestConfigurationsForTestCases);
            if (testCases == null)
            {
                SyncServiceTrace.D(Resources.TestCasesForTestPlanNotExist);
                throw new ArgumentNullException("testCases");
            }
            if (testPlan == null)
            {
                SyncServiceTrace.D(Resources.TestPlanNotExists);
                throw new ArgumentNullException("testPlan");
            }

            var testConfigurations = AvailableTestConfigurations;
            var retValue = new List<ITfsTestConfigurationDetail>();
            var testSuiteIds = new List<int>();

            if (tfsTestSuite != null)
            {
                testSuiteIds = GetAllTestSuiteIds(tfsTestSuite);
            }

            foreach (var testConfiguration in testConfigurations)
            {
                if (TestConfigurationIsAssigned(testPlan, testCases, testConfiguration, mostRecentTestResults, allConfigurations, testSuiteIds))
                {
                    retValue.Add(GetTestConfigurationDetail(testConfiguration));
                }
            }

            return retValue;
        }

        private List<int> GetAllTestSuiteIds(ITfsTestSuite testSuite)
        {
            var testSuiteIds = new List<int>();

            testSuiteIds.Add(testSuite.Id);

            foreach (var ts in testSuite.TestSuites)
            {
                var testPointIdsForCurrentSuite = GetAllTestSuiteIds(ts);
                testSuiteIds.AddRange(testPointIdsForCurrentSuite);

            }

            return testSuiteIds;

        }


        /// <summary>
        /// Get all Test Points for a multiple test cases
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="testCases"></param>
        /// <returns></returns>
        public IEnumerable<ITfsPropertyValueProvider> GetTestPointsForTestCases(ITfsTestPlan testPlan, IList<ITfsTestCaseDetail> testCases)
        {
            IList<ITfsTestPointDetail> retValue = new List<ITfsTestPointDetail>();
            SyncServiceTrace.D(Resources.GetTestPointsForTestCases);
            var tfsTestPlan = (ITestPlan)GetTestPlanDetail(testPlan).AssociatedObject;
            foreach (var tfsTestCaseDetail in testCases)
            {
                //Get all test points
                ITestPointCollection testPointCollection = tfsTestPlan.QueryTestPoints($"SELECT * FROM TestPoint WHERE TestCaseId = {tfsTestCaseDetail.Id}");

                foreach (var testPoint in testPointCollection)
                {
                    var tfsTestPoint = new TfsTestPoint(testPoint);
                    retValue.Add(new TfsTestPointDetail(tfsTestPoint));
                }

            }
            SyncServiceTrace.D(Resources.NumberOfTestPoints + retValue.Count());
            return retValue;
        }

        /// <summary>
        /// Get all test points for a testplan
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="boundElement"></param>
        /// <returns></returns>
        public IList<ITfsTestPointDetail> GetTestPointsForTestPlan(ITfsTestPlan testPlan, ITfsPropertyValueProvider boundElement)
        {
            SyncServiceTrace.D(Resources.GetTestPointsForTestPlan);
            var tfsTestPlan = (ITestPlan)GetTestPlanDetail(testPlan).AssociatedObject;
            IList<ITfsTestPointDetail> retValue = new List<ITfsTestPointDetail>();
            var query = "SELECT * FROM TestPoint";
            ITestPointCollection testPointCollection = tfsTestPlan.QueryTestPoints(query);

            foreach (var testPoint in testPointCollection)
            {
                var tfsTestPoint = new TfsTestPoint(testPoint);
                retValue.Add(new TfsTestPointDetail(tfsTestPoint));
            }
            SyncServiceTrace.D(Resources.NumberOfTestPoints + retValue.Count());
            return retValue;
        }

        public IList<ITfsTestPointDetail> GetTestPointsForTestSuite(ITfsTestPlan testPlan, ITfsPropertyValueProvider boundElement)
        {
            SyncServiceTrace.D(Resources.GetTestPointsForTestSuite);
            var tfsTestPlan = (ITestPlan)GetTestPlanDetail(testPlan).AssociatedObject;
            var testSuite = boundElement as TfsTestSuiteDetail;
            IList<ITfsTestPointDetail> retValue = new List<ITfsTestPointDetail>();

            if (testSuite != null)
            {
                var query = "SELECT * FROM TestPoint";
                ITestPointCollection testPointCollection = tfsTestPlan.QueryTestPoints(query);

                var filteredResult = testPointCollection.Where(x => x.SuiteId == testSuite.TestSuite.Id).ToList();
                foreach (var testPoint in filteredResult)
                {
                    TfsTestPoint tfsTestPoint = new TfsTestPoint(testPoint);
                    retValue.Add(new TfsTestPointDetail(tfsTestPoint));
                }
            }
            SyncServiceTrace.D(Resources.NumberOfTestPoints + retValue.Count());
            return retValue;
        }

        /// <summary>
        /// Determine if a test configuration is assigned to a test case by checking the corresponding testpoints of zhe given testcases
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="testCases"></param>
        /// <param name="testConfiguration"></param>
        /// <param name="latest"></param>
        /// <param name="allConfigurations"></param>
        /// <param name="testSuiteIds"></param>
        /// <returns></returns>
        private bool TestConfigurationIsAssigned(ITfsTestPlan testPlan, IList<ITfsTestCaseDetail> testCases, ITfsTestConfiguration testConfiguration, bool isLatest, bool allConfigurations, List<int> testSuiteIds)
        {
            //Get the test plan
            var tfsTestPlan = (ITestPlan)GetTestPlanDetail(testPlan).AssociatedObject;

            //Gather all test points
            var allTestPoints = new List<ITestPoint>();
            //ITestPointCollection allTestPoints;M
            foreach (var tfsTestCaseDetail in testCases)
            {
                //Get all test points
                allTestPoints.AddRange(tfsTestPlan.QueryTestPoints($"SELECT * FROM TestPoint WHERE TestCaseId = {tfsTestCaseDetail.Id}"));
            }

            //Filter for the relevant test points
            var testPointCollection = new List<ITestPoint>();

            //Filter for the relevant test points
            if (isLatest)
            {
                //if there are no recent results return, else filter the list
                if (allTestPoints.Any(x => x.MostRecentResult != null))
                {

                    if (allConfigurations)
                    {
                        testPointCollection.AddRange(allTestPoints.Where(x => x.MostRecentResult != null));
                    }
                    else
                    {
                        //var value = allTestPoints.Where(x => x.MostRecentResult != null).OrderByDescending(d => d.MostRecentResult.DateCreated).First();
                        //testPointCollection.Add(value);
                        testPointCollection.Add(allTestPoints.Where(x => x.MostRecentResult != null).OrderByDescending(d => d.MostRecentResult.DateCreated).First());
                        //testPointCollection.Add(allTestPoints.Where(x => x.MostRecentResult != null).OrderByDescending(d => d.MostRecentResult.DateCreated).First());
                    }
                }
            }
            else
            {
                testPointCollection.AddRange(allTestPoints);
            }

            //Check if it is assigned
            if (testPointCollection.Any(x => x.ConfigurationId == testConfiguration.Id))
            {

                if (testSuiteIds != null && testSuiteIds.Count > 0 && testPointCollection.Where(x => x.ConfigurationId == testConfiguration.Id).Any(x => testSuiteIds.Contains(x.SuiteId)))
                {
                    return true;
                }

                if (testSuiteIds != null && testSuiteIds.Count == 0 && testPointCollection.Any(x => x.ConfigurationId == testConfiguration.Id))
                {
                    return true;
                }
            }

            //if (latest)
            //{
            //    if (testPointCollection.Any(x => x.MostRecentResult != null))
            //    {


            //        if (allConfigurations)
            //        {
            //            if (!testPointCollection.Where(x => x.MostRecentResult != null).Any(x => x.ConfigurationId == testConfiguration.Id))
            //            {
            //                return false;
            //            }
            //        }
            //        else
            //        {
            //            if (!(testPointCollection.Where(x => x.MostRecentResult != null).OrderByDescending(d => d.MostRecentResult.DateCreated).First().ConfigurationId == testConfiguration.Id))
            //            {
            //                return false;
            //            }
            //        }
            //    }


            //}

            return false;
        }

        //if (latest)
        //    {
        //        if (testPointCollection.Any(x => x.MostRecentResult != null))
        //        {

        //        if (allConfigurations)
        //        {

        //            if (testPointCollection.Where(x => x.MostRecentResult != null).Any(x => x.ConfigurationId ==  testConfiguration.Id))
        //            {
        //                if (testSuiteIds != null && testSuiteIds.Count > 0 && testPointCollection.Where(x => x.ConfigurationId ==  testConfiguration.Id).Any(x => testSuiteIds.Contains(x.SuiteId)))
        //                {
        //                    return true;
        //                }

        //                if (testSuiteIds != null && testSuiteIds.Count == 0 && testPointCollection.Any(x => x.ConfigurationId == testConfiguration.Id))
        //                {
        //                    return true;
        //                }
        //            }
        //        }
        //        else
        //        {
        //            if (testPointCollection.Where(x => x.MostRecentResult != null).OrderByDescending(d => d.MostRecentResult.DateCreated).First().ConfigurationId == testConfiguration.Id)
        //        {
        //            if (testSuiteIds != null && testSuiteIds.Count > 0 && testPointCollection.Where(x => x.ConfigurationId == testConfiguration.Id).Any(x => testSuiteIds.Contains(x.SuiteId)))
        //            {
        //                return true;
        //            }

        //            if (testSuiteIds != null && testSuiteIds.Count == 0 && testPointCollection.Any(x => x.ConfigurationId == testConfiguration.Id))
        //            {
        //                return true;
        //            }
        //        }
        //        }

        //    }
        //    }
        //    else
        //    {

        //    }

        //}


        /// <summary>
        /// The method finds and returns the required test plan.
        /// </summary>
        /// <param name="testPlan"><see cref="ITfsTestPlan"/> specifies test plan which should be returned.</param>
        /// <returns>Required test plan.</returns>
        public ITfsTestPlanDetail GetTestPlanDetail(ITfsTestPlan testPlan)
        {
            if (testPlan == null)
            {
                return null;
            }
            return new TfsTestPlanDetail(testPlan as TfsTestPlan);
        }

        /// <summary>
        /// The method finds and returns the required test plan.
        /// </summary>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> specifies child test suite from which is the test plan to determine.</param>
        /// <returns>Required test plan.</returns>
        public ITfsTestPlanDetail GetTestPlanDetail(ITfsTestSuite testSuite)
        {
            var suite = testSuite as TfsTestSuite;
            if (suite == null)
            {
                return null;
            }
            ITestPlan testPlan = null;
            // 2016-07-18/TRU: Why not suite.AssotiatedPlan???
            if (suite.OriginalStaticTestSuite != null)
            {
                testPlan = suite.OriginalStaticTestSuite.Plan;
            }
            else if (suite.OriginalDynamicTestSuite != null)
            {
                testPlan = suite.OriginalDynamicTestSuite.Plan;
            }
            else if (suite.OriginalRequirementTestSuite != null)
            {
                testPlan = suite.OriginalRequirementTestSuite.Plan;
            }
            if (testPlan == null)
            {
                return null;
            }
            return new TfsTestPlanDetail(new TfsTestPlan(testPlan));
        }

        /// <summary>
        /// The method finds and returns the required test suite.
        /// </summary>
        /// <param name="testSuite"><see cref="ITfsTestSuite"/> specifies test suite which should be returned.</param>
        /// <returns>Required test suite.</returns>
        public ITfsTestSuiteDetail GetTestSuiteDetail(ITfsTestSuite testSuite)
        {
            if (testSuite == null)
            {
                throw new ArgumentNullException("testSuite");
            }
            return new TfsTestSuiteDetail(testSuite as TfsTestSuite);
        }

        /// <summary>
        /// The method finds and returns the required test configuration.
        /// </summary>
        /// <param name="testConfiguration"><see cref="ITfsTestConfiguration"/> specifies test configuration which should be returned.</param>
        /// <returns>Required test configuration.</returns>
        public ITfsTestConfigurationDetail GetTestConfigurationDetail(ITfsTestConfiguration testConfiguration)
        {
            if (testConfiguration == null)
            {
                SyncServiceTrace.D(Resources.TestConfigurationNotExists);
                throw new ArgumentNullException("testConfiguration");
            }
            return new TfsTestConfigurationDetail(testConfiguration as TfsTestConfiguration);
        }

        /// <summary>
        /// The method estimates all test cases that belongs to the given test plan.
        /// </summary>
        /// <param name="testPlan">Test plan whose test cases should be determined.</param>
        /// <param name="expandSharedSteps">Determines whether SharedSteps shall be shown expander or collapsed.</param>
        /// <returns>Determined test cases of given test plan.</returns>
        public IList<ITfsTestCaseDetail> GetTestCases(ITfsTestPlan testPlan, bool expandSharedSteps)
        {
            if (testPlan == null)
            {
                SyncServiceTrace.D(Resources.TestPlanNotExists);
                throw new ArgumentNullException("testPlan");
            }
            var retValue = new List<ITfsTestCaseDetail>();
            var plan = GetTestPlanDetail(testPlan) as TfsTestPlanDetail;
            if (plan == null)
            {
                return retValue;
            }
            var newTestSuite = new TfsTestSuite(plan.OriginalTestPlan.RootSuite, testPlan);
            var suiteDetail = GetTestSuiteDetail(newTestSuite) as TfsTestSuiteDetail;
            if (suiteDetail == null)
            {
                return retValue;
            }
            if (suiteDetail.OriginalStaticTestSuite != null)
            {
                foreach (var testCase in suiteDetail.OriginalStaticTestSuite.AllTestCases)
                {
                    var element = retValue.Find(
                        b => testCase.WorkItem != null && b.WorkItemId == testCase.WorkItem.Id);
                    if (element == null)
                    {
                        // Not in list yet
                        retValue.Add(new TfsTestCaseDetail(new TfsTestCase(testCase), expandSharedSteps));
                    }
                }
            }

            if (suiteDetail.OriginalDynamicTestSuite != null)
            {
                foreach (var testCase in suiteDetail.OriginalDynamicTestSuite.AllTestCases)
                {
                    var element = retValue.Find(
                        b => testCase.WorkItem != null && b.WorkItemId == testCase.WorkItem.Id);
                    if (element == null)
                    {
                        // Not in list yet
                        retValue.Add(new TfsTestCaseDetail(new TfsTestCase(testCase), expandSharedSteps));
                    }
                }
            }
            return retValue;
        }

        /// <summary>
        /// The method estimates all test cases that belongs to the given test suite and all dependent test suites.
        /// </summary>
        /// <param name="testSuite">Test suite whose test cases should be determined.</param>
        /// <param name="expandSharedSteps">Determines whether SharedSteps shall be shown expander or collapsed.</param>
        /// <returns>Determined test cases of given test suite.</returns>
        public IList<ITfsTestCaseDetail> GetAllTestCases(ITfsTestSuite testSuite, bool expandSharedSteps)
        {
            SyncServiceTrace.D(Resources.GetAllTestCasesBasedOnTestSuite);
            if (testSuite == null)
            {
                SyncServiceTrace.D(Resources.TestSuiteNotExist);
                throw new ArgumentNullException("testSuite");
            }
            var retValue = new List<ITfsTestCaseDetail>();
            var suiteDetail = GetTestSuiteDetail(testSuite) as TfsTestSuiteDetail;
            if (suiteDetail == null)
            {
                return retValue;
            }
            if (suiteDetail.OriginalStaticTestSuite != null)
            {
                foreach (var testCase in suiteDetail.OriginalStaticTestSuite.AllTestCases)
                {
                    var element = retValue.Find(
                        b => testCase.WorkItem != null && b.WorkItemId == testCase.WorkItem.Id);
                    if (element == null)
                    {
                        // Not in list yet
                        retValue.Add(new TfsTestCaseDetail(new TfsTestCase(testCase, testSuite), expandSharedSteps));
                    }
                }
            }

            if (suiteDetail.OriginalDynamicTestSuite != null)
            {
                foreach (var testCase in suiteDetail.OriginalDynamicTestSuite.AllTestCases)
                {
                    var element = retValue.Find(
                        b => testCase.WorkItem != null && b.WorkItemId == testCase.WorkItem.Id);
                    if (element == null)
                    {
                        // Not in list yet
                        retValue.Add(new TfsTestCaseDetail(new TfsTestCase(testCase, testSuite), expandSharedSteps));
                    }
                }
            }

            if (suiteDetail.OriginalRequirementTestSuite != null)
            {
                foreach (var testCase in suiteDetail.OriginalRequirementTestSuite.AllTestCases)
                {
                    var element = retValue.Find(
                        b => testCase.WorkItem != null && b.WorkItemId == testCase.WorkItem.Id);
                    if (element == null)
                    {
                        // Not in list yet
                        retValue.Add(new TfsTestCaseDetail(new TfsTestCase(testCase, testSuite), expandSharedSteps));
                    }
                }
            }

            return retValue;
        }

        /// <summary>
        /// The method estimates all test cases that belongs to the given test suite.
        /// </summary>
        /// <param name="testSuite">Test suite whose test cases should be determined.</param>
        /// <param name="expandSharedSteps">Determines whether SharedSteps shall be shown expander or collapsed.</param>
        /// <returns>Determined test cases of given test suite.</returns>
        public IList<ITfsTestCaseDetail> GetTestCases(ITfsTestSuite testSuite, bool expandSharedSteps)
        {
            if (testSuite == null)
            {
                SyncServiceTrace.D(Resources.TestSuiteNotExist);
                throw new ArgumentNullException("testSuite");
            }
            var retValue = new List<ITfsTestCaseDetail>();
            var suiteDetail = GetTestSuiteDetail(testSuite) as TfsTestSuiteDetail;
            if (suiteDetail == null)
            {
                return retValue;
            }
            if (suiteDetail.OriginalStaticTestSuite != null)
            {
                foreach (var testCase in suiteDetail.OriginalStaticTestSuite.TestCases)
                {
                    var element = retValue.Find(
                        b => testCase.TestCase.WorkItem != null && b.WorkItemId == testCase.TestCase.WorkItem.Id);
                    if (element == null)
                    {
                        // Not in list yet
                        retValue.Add(new TfsTestCaseDetail(new TfsTestCase(testCase.TestCase, testSuite), expandSharedSteps));
                    }
                }
            }
            if (suiteDetail.OriginalDynamicTestSuite != null)
            {
                foreach (var testCase in suiteDetail.OriginalDynamicTestSuite.TestCases)
                {
                    var element = retValue.Find(
                        b => testCase.TestCase.WorkItem != null && b.WorkItemId == testCase.TestCase.WorkItem.Id);
                    if (element == null)
                    {
                        // Not in list yet
                        retValue.Add(new TfsTestCaseDetail(new TfsTestCase(testCase.TestCase, testSuite), expandSharedSteps));
                    }
                }
            }

            if (suiteDetail.OriginalRequirementTestSuite != null)
            {
                foreach (var testCase in suiteDetail.OriginalRequirementTestSuite.TestCases)
                {
                    var element = retValue.Find(
                        b => testCase.TestCase.WorkItem != null && b.WorkItemId == testCase.TestCase.WorkItem.Id);
                    if (element == null)
                    {
                        // Not in list yet
                        retValue.Add(new TfsTestCaseDetail(new TfsTestCase(testCase.TestCase, testSuite), expandSharedSteps));
                    }
                }
            }

            //TODO MAYBE throw exception of no original test cases could be found check this (fore each distinguished test case type
            return retValue;
        }

        /// <summary>
        /// The method determines all test results that belongs to the given test case.
        /// </summary>
        /// <param name="testPlan">Test plan whose test results should be determined.</param>
        /// <param name="testCase">Test case whose test results should be determined.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        ///     If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        ///     If <c>null</c>, the build is no matter.</param>
        /// <param name="testSuite"></param>
        /// <param name="hierarchical">Determines if testresults for an testsuite should be searched hierarchical</param>
        /// <returns>Determined test results of given test case.</returns>
        public IList<ITfsTestResultDetail> GetTestResults(ITfsTestPlan testPlan, ITfsTestCase testCase, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild, ITfsTestSuite testSuite, bool hierarchical, out Dictionary<int, int> lastTestRunPerConfig)
        {
            SyncServiceTrace.D(Resources.GetAllTestResultsBasedOnTestCase);
            Guard.ThrowOnArgumentNull(testCase, "testCase");
            //Guard.ThrowOnArgumentNull(testPlan, "testPlan");

            // In some cases is TestCaseResult.BuildNumber 'null', but result is joined with build.
            // In this case is the connection over TestRun provided.


            //if a testSuite is provided, use it to filter for all relevant testresults for this suite. Otherwise this will lead to a wrong mapping of  testresults with testsuites

            // var testPointsIds = new Dictionary<int, int>();
            var testPointsIds = new List<int>();
            lastTestRunPerConfig = new Dictionary<int, int>();

            if (testSuite == null)
            {
                testSuite = testPlan.RootTestSuite;
            }

            if (testSuite != null)
            {
                testPointsIds = GetTestPointIds(testPlan, testSuite, testCase, hierarchical, ref lastTestRunPerConfig);
            }

            var retValue = new List<ITfsTestResultDetail>();

            // Select test results
            var testResults =
                new List<ITestCaseResult>(
                    TestManagement.TestResults.Query($"SELECT * FROM TestResult WHERE TestCaseId = {testCase.Id}"));

            // Filter for test configuration
            if (testConfiguration != null && testConfiguration.Id != -1)
            {
                testResults = testResults.FindAll(testResult => testResult.TestConfigurationId == testConfiguration.Id);
            }

            // Select test runs for build
            List<ITestRun> buildTestRuns = null;
            if (serverBuild != null)
            {
                buildTestRuns = new List<ITestRun>(TestManagement.TestRuns.ByBuild(serverBuild.Uri));
            }

            // Iterate selected test results
            foreach (var testResult in testResults)
            {

                var associatedTestRun = TestManagement.TestRuns.Find(testResult.TestRunId);
                if (associatedTestRun == null ||
                     (testPlan != null && associatedTestRun.TestPlanId != testPlan.Id) ||
                     (!testPointsIds.Contains(testResult.TestPointId) && testPointsIds.Count != 0))
                {
                    continue;
                }

                var lastTestRunId = 0;
                if (lastTestRunPerConfig.Count != 0)
                {
                    lastTestRunId = lastTestRunPerConfig[testResult.TestConfigurationId];
                }

                if (serverBuild == null)
                {
                    retValue.Add(new TfsTestResultDetail(new TfsTestResult(testResult, testCase.Id), lastTestRunId));
                }
                else if (testResult.BuildNumber == serverBuild.BuildNumber)
                {
                    retValue.Add(new TfsTestResultDetail(new TfsTestResult(testResult, testCase.Id), lastTestRunId));
                }
                else if (buildTestRuns.Any(testRun => testResult.TestRunId == testRun.Id))
                {
                    retValue.Add(new TfsTestResultDetail(new TfsTestResult(testResult, testCase.Id), lastTestRunId));
                }
            }
            if (testPlan != null)
            {
                SyncServiceTrace.W(string.Format(Resources.NumberOfTestResultForTestPlan, retValue.Count, testPlan.Id));
            }
            else
            {
                SyncServiceTrace.W(string.Format(Resources.NumberOfTestResults, retValue.Count));
            }
            // No test result for the filter found.
            return retValue;
        }


        /// <summary>
        /// Get a dictionary of Test Point IDs with latest Test Run ID.
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="testSuite"></param>
        /// <param name="testCase"></param>
        /// <param name="isHierarchical"></param>
        /// <returns>Returns a dictionary having TestPointId as key and MostRecentRunId a value.</returns>
        private List<int> GetTestPointIds(ITfsTestPlan testPlan, ITfsTestSuite testSuite, ITfsTestCase testCase, bool isHierarchical, ref Dictionary<int, int> latestTestRunPerConfig)
        {
            SyncServiceTrace.D(Resources.GetTestPointIds);
            var testPointsIds = new List<int>();
            if (latestTestRunPerConfig == null)
            {
                latestTestRunPerConfig = new Dictionary<int, int>();
            }

            var tfsTestPlan = (ITestPlan)GetTestPlanDetail(testPlan).AssociatedObject;

            //Query for the testpoints
            var queryForTestPointsForSpecificTestSuite = $"SELECT * FROM TestPoint WHERE SuiteId = {testSuite.Id} and TestCaseId = {testCase.Id}";
            var testPoints = tfsTestPlan.QueryTestPoints(queryForTestPointsForSpecificTestSuite);

            if (testPoints != null)
            {
                SyncServiceTrace.D(Resources.NumberOfTestPoints + testPoints.Count);
            }

            if (testPoints.Count != 0)
            {
                // testPointsIds = (from testPoint in testPoints select testPoint).ToDictionary(p => p.Id, p => p.MostRecentRunId);
                testPointsIds = (from testPoint in testPoints select testPoint.Id).ToList();
                // testPointsIds = (from testPoint in testPoints select testPoint).ToDictionary(p => p.Id, new Dictionary<int, int>(p => p.ConfigurationId, p => p.MostRecentRunId));
                var latestTestRunPerConfigTemp = (from testPoint in testPoints select testPoint).GroupBy(p => p.ConfigurationId, p => p.MostRecentRunId).ToDictionary(group => group.Key, group => group.FirstOrDefault());
                foreach (var item in latestTestRunPerConfigTemp.Keys)
                {
                    latestTestRunPerConfig.Add(item, latestTestRunPerConfigTemp[item]);
                }
            }

            if (isHierarchical)
            {
                foreach (var ts in testSuite.TestSuites)
                {
                    // ReSharper disable once ConditionIsAlwaysTrueOrFalse
                    var testPointIdsForCurrentSuite = GetTestPointIds(testPlan, ts, testCase, isHierarchical, ref latestTestRunPerConfig);
                    testPointsIds.AddRange(testPointIdsForCurrentSuite);
                }
            }

            return testPointsIds;
        }

        /// <summary>
        /// Wrapper to get the latest test results. The configuration is not used here
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="testCase"></param>
        /// <param name="testConfiguration"></param>
        /// <param name="serverBuild"></param>
        /// <param name="testSuite"></param>
        /// <param name="hierarchical">Determines if testresults for an testsuite should be searched hierarchical</param>
        /// <returns></returns>
        public IList<ITfsTestResultDetail> GetLatestTestResults(ITfsTestPlan testPlan, ITfsTestCase testCase, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild, ITfsTestSuite testSuite, bool hierarchical)
        {
            var testResults = new List<ITfsTestResultDetail>();
            Dictionary<int, int> lastTestRunPerConfig;
            var allTestResultsSortedByLatest = GetTestResults(testPlan, testCase, testConfiguration, serverBuild, testSuite, hierarchical, out lastTestRunPerConfig).OrderBy(x => x.DateCreated);

            if (allTestResultsSortedByLatest.Any())
            {
                // latest run ID has the same value for each entry
                // var latestRunId = allTestResultsSortedByLatest.First().LatestTestRunId;

                var latestRunId = lastTestRunPerConfig.Values.FirstOrDefault();
                if (testConfiguration != null)
                {
                    latestRunId = lastTestRunPerConfig[testConfiguration.Id];
                }
                testResults = allTestResultsSortedByLatest.Where(x => ((x.AssociatedObject as ITestCaseResult).TestRunId == latestRunId)).ToList();
            }

            return testResults;

        }

        /// <summary>
        /// Wrapper to get the latest test results for all selected configurations
        /// </summary>
        /// <param name="testPlan"></param>
        /// <param name="testCase"></param>
        /// <param name="testConfiguration"></param>
        /// <param name="serverBuild"></param>
        /// <param name="testSuite"></param>
        /// <param name="hierarchical">Determines if testresults for an testsuite should be searched hierarchical</param>
        /// <returns></returns>
        public IList<ITfsTestResultDetail> GetLatestTestResultsForAllSelectedConfigurations(ITfsTestPlan testPlan, ITfsTestCase testCase, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild, ITfsTestSuite testSuite, bool hierarchical)
        {
            var testResults = new List<ITfsTestResultDetail>();
            Dictionary<int, int> lastTestRunPerConfig;
            var allTestResultsSortedByLatest = GetTestResults(testPlan, testCase, testConfiguration, serverBuild, testSuite, hierarchical, out lastTestRunPerConfig); //.GroupBy(x => x.TestConfigurationId).Select(y => y.OrderBy(x => x.DateCreated).Last());

            if (allTestResultsSortedByLatest.Any())
            {
                // latest run ID has the same value for each entry
                // var latestRunId = allTestResultsSortedByLatest.First().LatestTestRunId;
                // int latestRunId;

                if (testConfiguration != null && testConfiguration.Id != -1)
                {
                    // latestRunId = lastTestRunPerConfig[testConfiguration.Id];
                    lastTestRunPerConfig = lastTestRunPerConfig.Where(c => (c.Key == testConfiguration.Id)).ToDictionary(c => c.Key, c => c.Value);
                }
                else
                {
                    // latestRunId = lastTestRunPerConfig.Values.FirstOrDefault();
                }

                foreach (var cfg in lastTestRunPerConfig)
                {
                    var tmpTestResults = allTestResultsSortedByLatest.Where(x => ((x.AssociatedObject as ITestCaseResult).TestRunId == cfg.Value)).ToList();
                    testResults.AddRange(tmpTestResults);
                }

                // testResults = allTestResultsSortedByLatest.Where(x => ((x.AssociatedObject as ITestCaseResult).TestRunId == latestRunId)).ToList();
            }

            return testResults;
        }

        /// <summary>
        /// The method determines if at least one test result exists for given test case.
        /// </summary>
        /// <param name="testPlan">Test plan whose test results should be determined.</param>
        /// <param name="testCase">Test case to determine if at least one test result exists.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        /// If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        /// If <c>null</c>, the build is no matter.</param>
        /// <returns><c>true</c> for the test cases exists at least one test result. Otherwise <c>false</c>.</returns>
        public bool TestResultExists(ITfsTestPlan testPlan, ITfsTestSuite testSuite, ITfsTestCaseDetail testCase, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild)
        {
            Guard.ThrowOnArgumentNull(testPlan, "testPlan");
            Guard.ThrowOnArgumentNull(testCase, "testCase");

            // In some cases is TestCaseResult.BuildNumber 'null', but result is joined with build.
            // In this case is the connection over TestRun provided.

            //If a test suite is provided, use it to filter for all relevant test results for this suite. Otherwise this will lead to a wrong mapping of test results with test suites
            var testPointsIds = new List<int>();
            var lastTestRunPerConfig = new Dictionary<int, int>();

            if (testSuite == null)
            {
                testSuite = testPlan.RootTestSuite;
            }

            if (testSuite != null)
            {
                testPointsIds = GetTestPointIds(testPlan, testSuite, testCase.TestCase, false, ref lastTestRunPerConfig);
            }

            // Select test results
            var preTestResults = new List<ITestCaseResult>(TestManagement.TestResults.Query($"SELECT * FROM TestResult WHERE TestCaseId = {testCase.Id}"));

            var realTestResults = new List<ITestCaseResult>();
            foreach (var testResult in preTestResults)
            {
                var associatedTestRun = TestManagement.TestRuns.Find(testResult.TestRunId);
                // third condition - Skip test result which doesn't belong this test suite
                if (associatedTestRun == null || associatedTestRun.TestPlanId != testPlan.Id || (!testPointsIds.Contains(testResult.TestPointId) && testPointsIds.Count != 0))
                {
                    continue;
                }
                realTestResults.Add(testResult);
            }
            if (realTestResults.Count == 0)
            {
                return false;
            }

            // Filter for test configuration
            if (testConfiguration != null)
            {
                realTestResults = realTestResults.FindAll(testResult => testResult.TestConfigurationId == testConfiguration.Id);
            }

            // If no server build and at least one test result exists,
            if (serverBuild == null)
            {
                return realTestResults.Count > 0;
            }

            // If any test result has the appropriate build number
            if (realTestResults.Any(testResult => testResult.BuildNumber == serverBuild.BuildNumber))
            {
                return true;
            }

            // Select test runs for build
            var buildTestRuns = new List<ITestRun>(TestManagement.TestRuns.ByBuild(serverBuild.Uri));

            // Iterate through all test results
            if (realTestResults.Any(testResult => buildTestRuns.Any(testRun => testResult.TestRunId == testRun.Id)))
            {
                return true;
            }

            // nothing found
            return false;
        }

        /// <summary>
        /// The method determines if at least one test result exists for given test cases.
        /// </summary>
        /// <param name="testPlan">Test plan whose test results should be determined.</param>
        /// <param name="testCases">Test cases to determine if at least one test result exists.</param>
        /// <param name="testConfiguration">Filter for the operation. If defined, check test results only for this configuration.
        /// If <c>null</c>, the configuration is no matter.</param>
        /// <param name="serverBuild">Filter for the operation. If defined, check test results only for this build.
        /// If <c>null</c>, the build is no matter.</param>
        /// <returns><c>true</c> for the test cases exists at least one test result. Otherwise <c>false</c>.</returns>
        public bool TestResultExists(ITfsTestPlan testPlan, ITfsTestSuite testSuite, IEnumerable<ITfsTestCaseDetail> testCases, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild)
        {
            return testCases.Any(testCase => TestResultExists(testPlan, testSuite, testCase, testConfiguration, serverBuild));
        }

        /// <summary>
        /// The method determines the hyperlink for the build to the team foundation server web access.
        /// </summary>
        /// <param name="buildNumber">Build number of the build.</param>
        /// <returns>Required hyperlink.</returns>
        public Uri GetBuildViewerLink(string buildNumber)
        {
            var builds = GetAllBuilds();
            if (builds == null)
            {
                return null;
            }
            var build = builds.FirstOrDefault(b => b.BuildNumber == buildNumber);
            if (build == null)
            {
                return null;
            }
            var hyperlinkService = TfsTeamProjectCollection.GetService<TswaClientHyperlinkService>();
            return hyperlinkService.GetViewBuildDetailsUrl(build.Uri);
        }

        /// <summary>
        /// The method determines the hyperlink for the work item to the team foundation server web access editor of the work item.
        /// </summary>
        /// <param name="workItemArtifactLink">Artifact link of the work item.</param>
        /// <returns>Required hyperlink.</returns>
        public Uri GetWorkItemEditorLink(string workItemArtifactLink)
        {
            var hyperlinkService = TfsTeamProjectCollection.GetService<TswaClientHyperlinkService>();
            return hyperlinkService.GetArtifactViewerUrl(new Uri(workItemArtifactLink));
        }

        /// <summary>
        /// The method determines the hyperlink for the work item to the team foundation server web access viewer of the work item.
        /// </summary>
        /// <param name="workItemNumber">Number of the work item.</param>
        /// <param name="revision">Required revision of work item. If not defined, last revision is used.</param>
        /// <returns>Required hyperlink.</returns>
        public Uri GetWorkItemViewerLink(int workItemNumber, int revision)
        {
            var link = string.Empty;
            if (revision == -1)
                link = $"{TfsTeamProjectCollection.Uri.AbsoluteUri}/WorkItemTracking/workitem.aspx?artifactMoniker={workItemNumber}";
            else
                link = $"{TfsTeamProjectCollection.Uri.AbsoluteUri}/WorkItemTracking/workitem.aspx?artifactMoniker={workItemNumber}&Rev={revision}";
            return new Uri(link);
        }

        /// <summary>
        /// Gets the artifact link.
        /// Valid artifact types are: <code>Project, WorkItem, Query, VersionedItem, LatestItemVersion, Changeset, Shelveset, ShelvedItem and Build</code>.
        /// </summary>
        /// <param name="artifactLink">A valid team foundation server artifact Uri.</param>
        /// <returns>A viewer URL for a TFS artifact.</returns>
        public Uri GetArtifactLink(string artifactLink)
        {
            var hyperlinkService = TfsTeamProjectCollection.GetService<TswaClientHyperlinkService>();
            return hyperlinkService.GetArtifactViewerUrl(new Uri(artifactLink));
        }

        /// <summary>
        /// Gets the title of work item.
        /// </summary>
        /// <param name="workItemId">Id of the work item to get the title for.</param>
        /// <returns>Title of the work item if the work item found. Otherwise <c>null</c>.</returns>
        public string GetWorkItemTitle(int workItemId)
        {
            var queryString = $"SELECT * FROM WorkItems WHERE System.Id={workItemId} AND System.TeamProject='{ProjectName}'";
            var workItemCollection = WorkItemStore.Query(queryString);
            if (workItemCollection != null && workItemCollection.Count == 1)
            {
                return workItemCollection[0].Title;
            }
            return null;
        }

        /// <summary>
        /// The method expands the enumerable. Used at least for shared steps in test case actions.
        /// </summary>
        /// <param name="enumerable">Enumerable to expand.</param>
        /// <returns>Expanded enumerable.</returns>
        public IEnumerable ExpandEnumerable(IEnumerable enumerable)
        {
            var actions = enumerable as TestActionCollection;
            if (actions == null)
            {
                return enumerable;
            }

            var retList = new List<ITestStep>();
            foreach (var action in actions)
            {
                var testStep = action as ITestStep;
                var sharedSteps = action as ISharedStepReference;
                if (testStep != null)
                {
                    retList.Add(testStep);
                }
                else if (sharedSteps != null)
                {
                    var steps = TestManagement.SharedSteps.Query($"SELECT * FROM WorkItem WHERE Id = {sharedSteps.SharedStepId}");
                    retList.AddRange(
                        from sharedTestSteps in steps
                        from ITestStep sharedTestStep in sharedTestSteps.Actions
                        where sharedTestStep != null
                        select sharedTestStep);
                }
            }
            return retList;
        }

        /// <summary>
        /// Get all linked WorkItems of a specific Type and link for a Workitem with a specific id.
        /// This is a wrapper that will return all linked items
        /// </summary>
        /// <param name="workItemId">The Id of the WorkItem</param>
        /// <param name="workItemTypeCsv"></param>
        /// <param name="linkTypeCsv"></param>
        /// <param name="filterOption"></param>
        /// <returns></returns>
        public List<WorkItem> GetAllLinkedWorkItemsForWorkItemId(int workItemId, string workItemTypeCsv, string linkTypeCsv, IFilterOption filterOption)
        {

            List<string> workItemTypes = CsvParser.ParseCsv(workItemTypeCsv);
            List<string> linkTypes = CsvParser.ParseCsv(linkTypeCsv);
            var filter = new WorkItemLinkFilter();
            filter.FilterType = FilterType.Include;
            foreach (var linkType in linkTypes)
            {
                var singleFilter = new Filter();
                singleFilter.LinkType = linkType;

            }
            return GetAllLinkedWorkItemsForWorkItemId(workItemId, workItemTypes, filter, filterOption, null);


        }

        /// <summary>
        /// Get all linked WorkItems of a specific Type and link for a Workitem with a specific id.
        /// </summary>
        /// <param name="workItemId">The Id of the WorkItem</param>
        /// <param name="workItemTypes"></param>
        /// <param name="linkFilter"></param>
        /// <param name="filterOption"></param>
        /// <param name="returnList"></param>
        /// <returns></returns>
        private List<WorkItem> GetAllLinkedWorkItemsForWorkItemId(int workItemId, List<string> workItemTypes, IWorkItemLinkFilter linkFilter, IFilterOption filterOption, List<WorkItem> returnList)
        {

            if (returnList == null)
                returnList = new List<WorkItem>();

            var workItem = WorkItemStore.GetWorkItem(workItemId);

            //If TestCases are included, add the TestCases itself to list
            if (workItemTypes.Contains("Test Cases"))
            {
                returnList.Add(workItem);
            }

            var links = workItem.WorkItemLinks;
            foreach (WorkItemLink link in links)
            {
                //Check the immutable Name
                var name = link.LinkTypeEnd.LinkType.ForwardEnd.ImmutableName;

                //If exclude filter --> Continue (Skip) if the name is in the list
                //If include ---> Continue (Skip) if the name is not in the list
                if ((linkFilter != null && linkFilter.Filters.Count > 0)
                    && ((linkFilter.FilterType.Equals(FilterType.Exclude) && (linkFilter.Filters.Any(x => x.LinkType.Equals(name))))
                        || (linkFilter.FilterType.Equals(FilterType.Include) && !(linkFilter.Filters.Any(x => x.LinkType.Equals(name))))))
                {
                    continue;
                }

                var relatedWorkItem = WorkItemStore.GetWorkItem(link.TargetId);

                //Check the type of the linked workitem and add it to the return list
                var workItemTypeName = relatedWorkItem.Type.Name;

                if (!(workItemTypes.Contains(workItemTypeName)))
                {
                    continue;
                }


                if (filterOption != null)
                {
                    if (!(filterOption.Distinct && returnList.Any(x => x.Id.Equals(relatedWorkItem.Id))))
                    {
                        returnList.Add(relatedWorkItem);
                    }
                }
                else
                {
                    returnList.Add(relatedWorkItem);
                }

            }
            return returnList;
        }

        /// <summary>
        /// Get linked workitems for testcase.
        /// This is warpper for GetAllLinkedWorkItemsForWorkItemId
        /// </summary>
        /// <param name="testCases"></param>
        /// <param name="workItemTypes"></param>
        /// <param name="linkFilter"></param>
        /// <param name="filterOption"></param>
        /// <returns></returns>
        public List<WorkItem> GetAllLinkedWorkItemsForTestCases(IList<ITfsTestCaseDetail> testCases, List<string> workItemTypes, IWorkItemLinkFilter linkFilter, IFilterOption filterOption)
        {
            Guard.ThrowOnArgumentNull(testCases, "testCases");

            var returnList = new List<WorkItem>();
            foreach (var testCase in testCases)
            {
                var linkedWorkItems = GetAllLinkedWorkItemsForWorkItemId(testCase.Id, workItemTypes, linkFilter, filterOption, null);
                returnList.AddRange(linkedWorkItems);
            }
            return returnList;
        }

        /// <summary>
        /// This method gets all linked workitems for a test result
        /// </summary>
        /// <param name="testResult">The test result</param>
        /// <param name="workItemTypes">The work item types that should be returned</param>
        /// <param name="linkFilter">the link filters</param>
        /// <param name="filterOption">the filter options</param>
        /// <returns></returns>
        public List<WorkItem> GetAllLinkedWorkItemsForTestResult(ITfsTestResultDetail testResult, List<string> workItemTypes, IWorkItemLinkFilter linkFilter, IFilterOption filterOption)
        {
            Guard.ThrowOnArgumentNull(testResult, "testResult");
            SyncServiceTrace.W(string.Format(Resources.SearchAllLinkedWorkItemsForTestResult, testResult.TestResult.TestCaseId));

            var returnList = new List<WorkItem>();
            var testCaseResult = (ITestCaseResult)testResult.AssociatedObject;

            var associatedWorkItems = testCaseResult.QueryAssociatedWorkItems();
            SyncServiceTrace.W(string.Format(Resources.NumberOfAssociatedWorkItems, associatedWorkItems.Count()));

            foreach (int id in associatedWorkItems)
            {
                var relatedWorkItem = WorkItemStore.GetWorkItem(id);

                var links = relatedWorkItem.Links;
                foreach (var link in links)
                {
                    var linkName = string.Empty;

                    if (link is WorkItemLink)
                    {
                        //Check the immutable Name
                        linkName = ((WorkItemLink)link).LinkTypeEnd.LinkType.ForwardEnd.ImmutableName;

                    }
                    if (link is RelatedLink)
                    {
                        var rlink = (RelatedLink)link;
                        linkName = rlink.LinkTypeEnd.LinkType.ForwardEnd.ImmutableName;
                    }

                    if (link is ExternalLink)
                    {
                        var rlink = (ExternalLink)link;
                        linkName = rlink.ArtifactLinkType.Name;
                    }
                    if (linkFilter != null && linkFilter.Filters.Count > 0)
                    {
                        //First condition - If exclude filter --> Continue (Skip) if the name is in the list
                        //Second condition - If include ---> Continue (Skip) if the name is not in the list (ReSharper disable once RedundantJumpStatement)
                        if ((linkFilter.FilterType.Equals(FilterType.Exclude) && (linkFilter.Filters.Any(x => x.LinkType.Equals(linkName))))
                            || (linkFilter.FilterType.Equals(FilterType.Include) && !(linkFilter.Filters.Any(x => x.LinkType.Equals(linkName)))))
                        {
                            continue;
                        }
                    }
                }

                if (!(workItemTypes.Contains(relatedWorkItem.Type.Name)))
                {
                    continue;
                }

                if (filterOption != null)
                {
                    if (!(filterOption.Distinct && returnList.Any(x => x.Id.Equals(relatedWorkItem.Id))))
                    {
                        returnList.Add(relatedWorkItem);
                    }
                }
                else
                {
                    returnList.Add(relatedWorkItem);
                }

            }
            SyncServiceTrace.W(string.Format(Resources.NumberOfMatchedAssociatedWorkItems, returnList.Count()));

            return returnList;
        }

        /// <summary>
        /// Get all builds for a given TestCase. This method queries the associated TestResults and returns the associated bu
        /// </summary>
        /// <param name="tfsTestCaseDetail"></param>
        /// <returns></returns>
        public IEnumerable<IBuildDetail> GetAllBuildsForTestCase(ITfsTestCaseDetail tfsTestCaseDetail, bool reloadBuilds)
        {
            Guard.ThrowOnArgumentNull(tfsTestCaseDetail, "tfsTestCaseDetail");
            Guard.ThrowOnArgumentNull(tfsTestCaseDetail.TestCase.AssociatedTestSuiteDetail, "No testsuite for testcase present");

            if (reloadBuilds)
            {
                _builds = null;
            }

            if (Builds == null)
            {
                return null;
            }

            var returnList = new List<IBuildDetail>();
            Dictionary<int, int> lastTestRunPerConfig;
            //Get the test results
            IEnumerable<ITfsTestResultDetail> testResults = GetTestResults(tfsTestCaseDetail.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tfsTestCaseDetail.TestCase, null, null, null, false, out lastTestRunPerConfig);

            foreach (var tr in testResults)
            {
                var tcr = (ITestCaseResult)tr.AssociatedObject;
                if (tcr.BuildNumber == null)
                {
                    continue;
                }

                var build = Builds.Where(b => b.BuildNumber.Equals(tcr.BuildNumber)).FirstOrDefault<IBuildDetail>();
                if (!returnList.Any(x => x.BuildNumber.Equals(build.BuildNumber)))
                {
                    returnList.Add(build);
                }
            }
            return returnList;
        }

        /// <summary>
        /// This method queries builds for a specific test case
        /// </summary>
        /// <param name="tfsTestCaseDetail">The specific Test Case</param>
        /// <param name="latestTestResults">Filter for the opertation, if true only the builds for the latest test result are returned</param>
        /// <param name="latestTestResultsForAllConfigurations">Filter for the opertation, if true only the builds for the latest test result for all configurations are returned</param>
        /// <param name="testConfiguration"></param>
        /// <returns></returns>
        public IEnumerable<IBuildDetail> GetBuildsForTestCase(ITfsTestCaseDetail tfsTestCaseDetail, bool latestTestResults, bool latestTestResultsForAllConfigurations, ITfsTestConfiguration testConfiguration, bool reloadBuilds)
        {
            if (reloadBuilds)
            {
                _builds = null;
            }
            if (Builds == null)
            {
                return null;
            }

            SyncServiceTrace.W(string.Format(Resources.SearchBuildsForTestCase, tfsTestCaseDetail.TestCase.Id, latestTestResults.ToString(), latestTestResultsForAllConfigurations.ToString()));
            var returnList = new List<IBuildDetail>();

            //Get the test results
            IEnumerable<ITfsTestResultDetail> testResults;
            if (latestTestResults)
            {
                //Get the last testresult for each configuration
                if (latestTestResultsForAllConfigurations)
                {
                    testResults = GetLatestTestResultsForAllSelectedConfigurations(tfsTestCaseDetail.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tfsTestCaseDetail.TestCase, testConfiguration, null, tfsTestCaseDetail.TestCase.AssociatedTestSuiteDetail, true);
                }
                else
                {
                    testResults = GetLatestTestResults(tfsTestCaseDetail.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tfsTestCaseDetail.TestCase, null, null, tfsTestCaseDetail.TestCase.AssociatedTestSuiteDetail, true);
                }
            }
            else
            {
                Dictionary<int, int> lastTestRunPerConfig;
                testResults = GetTestResults(tfsTestCaseDetail.TestCase.AssociatedTestSuiteDetail.AssociatedTestPlan, tfsTestCaseDetail.TestCase, null, null, tfsTestCaseDetail.TestCase.AssociatedTestSuiteDetail, true, out lastTestRunPerConfig);
            }

            foreach (var testResult in testResults)
            {
                var testCaseResult = (ITestCaseResult)testResult.AssociatedObject;
                if (testCaseResult.BuildNumber == null)
                {
                    continue;
                }

                var build = Builds.Where(b => b.BuildNumber.Equals(testCaseResult.BuildNumber)).FirstOrDefault<IBuildDetail>();
                if (!returnList.Any(x => x.BuildNumber.Equals(build.BuildNumber)))
                {
                    SyncServiceTrace.W(string.Format(Resources.BuildIsLinkedTo, build.BuildNumber, testResult.TestResult.Id));
                    returnList.Add(build);
                }
            }
            return returnList;
        }

        #endregion Implementation of ITfsTestAdapter

        #region Private methods

        /// <summary>
        /// Recursively visits all nodes in sub trees while creating another tree
        /// with the same structure but different node types.
        /// </summary>
        private static void GetAreaIterationNodes(IAreaIterationNode parentNode, NodeCollection children, IAreaIterationNode rootNode)
        {
            if (children == null)
            {
                return;
            }

            foreach (Node child in children)
            {
                var childNode = new AreaIterationNode(child.Name, child.Path, child.Id.ToString(CultureInfo.InvariantCulture), rootNode);
                parentNode.Childs.Add(childNode);
                GetAreaIterationNodes(childNode, child.ChildNodes, rootNode);
            }
        }

        /// <summary>
        /// Gets the localized (server language) name of a structure type
        /// </summary>
        /// <param name="structureType">Name or constant from Microsoft.TeamFoundation.Proxy.StructureType</param>
        private string GetStructureName(string structureType)
        {
            var commonStructureService = TfsTeamProjectCollection.GetService<ICommonStructureService>();
            var projectInfo = commonStructureService.GetProjectFromName(ProjectName);
            var structures = commonStructureService.ListStructures(projectInfo.Uri);
            return structures.First(x => x.StructureType.Equals(structureType)).Name;
        }

        /// <summary>
        /// Method saves all changed work items to team foundation server.
        /// </summary>
        /// <returns>Occurred errors.</returns>
        private IList<ISaveError> InternalSave(IList<TfsWorkItem> workItems)
        {
            var workItemsToSave = new List<WorkItem>();
            var workItemsErrorResolver = new Dictionary<WorkItem, IWorkItem>();

            foreach (TfsWorkItem workItem in workItems)
            {
                if (workItem.WorkItem.IsDirty && !workItem.HasWordValidationError)
                {
                    workItemsToSave.Add(workItem.WorkItem);
                    workItemsErrorResolver.Add(workItem.WorkItem, workItem);
                }
            }

            if (0 != workItemsToSave.Count)
            {
                var errors = WorkItemStore.BatchSave(workItemsToSave.ToArray());
                IList<ISaveError> saveErrors = SaveTestCases(new List<ISaveError>());

                foreach (BatchSaveError saveError in errors)
                {
                    saveErrors.Add(new SaveError(workItemsErrorResolver[saveError.WorkItem], saveError.Exception));
                }
                return saveErrors;
            }

            return SaveTestCases(new List<ISaveError>());
        }

        /// <summary>
        /// Save test cases for work items. Do this after work items are saved
        /// </summary>
        private IList<ISaveError> SaveTestCases(IList<ISaveError> saveErrors)
        {
            foreach (TfsWorkItem workItem in WorkItems)
            {
                var testCaseField = workItem.Fields.FirstOrDefault(x => x is TfsTestCaseField) as TfsTestCaseField;

                if (testCaseField != null)
                {
                    testCaseField.SaveTestCaseValue();
                }
            }
            return saveErrors;
        }

        /// <summary>
        /// Load all work items to process.
        /// </summary>
        /// <returns>True on success, otherwise false.</returns>
        private void LoadWorkItems(int[] ids, IList<string> fields)
        {
            SyncServiceTrace.D(Resources.LoadWorkItems);
            if (_configuration == null)
            {
                SyncServiceTrace.D(Resources.ConfigurationNotExists);
                return;
            }
            WorkItems = new TfsWorkItemCollection();

            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.EnterProgressGroup(2, Resources.LoadWIs_Querying);
            progressService.DoTick();

            IList<WorkItem> queryWorkItems = new List<WorkItem>();
            if (_queryConfiguration == null)
            {
                queryWorkItems = LoadWorkItemsWithoutConfiguration(ids, fields);
            }
            else
            {
                switch (_queryConfiguration.ImportOption)
                {
                    case QueryImportOption.SavedQuery:
                        queryWorkItems = LoadWorkItemsFromSavedQuery(_queryConfiguration);
                        break;
                    case QueryImportOption.IDs:
                        queryWorkItems = LoadWorkItemsById(_queryConfiguration);
                        break;
                    case QueryImportOption.TitleContains:
                        queryWorkItems = LoadWorkItemsByTitle(_queryConfiguration);
                        break;
                }

                // Apply custom sorting unless we have a saved query that does not use linked work items
                if (_queryConfiguration.ImportOption != QueryImportOption.SavedQuery ||
                    _queryConfiguration.UseLinkedWorkItems)
                {
                    queryWorkItems = GetHierarchyWorkItemsByConfiguration(queryWorkItems, _queryConfiguration,
                                                                          WorkItemStore);
                }
            }

            progressService.DoTick();
            progressService.LeaveProgressGroup();

            // Create wrapper for each found work item. Assign either the
            // standard configuration or one set explicitly using a paramter
            foreach (var workItem in queryWorkItems)
            {
                var workItemToUse = workItem;
                var config = _configuration.GetWorkItemConfigurationExtended(workItem.Type.Name,
                    fieldName =>
                    {
                        if (string.IsNullOrEmpty(fieldName))
                        {
                            return null;
                        }

                        //Iterate through the tree and return the current level
                        //Do this only if the parameter "TypeOfHierachyRelationships" is specified in the mapping file
                        //Get The hierarchylevel by a second function
                        var nameOfRelationship = _configuration.TypeOfHierarchyRelationships;
                        if (fieldName.Equals(Constants.NameOfHierarchyLevelProperty))
                        {
                            //Return null if the parameter is not specified or empty
                            if (nameOfRelationship == null)
                            {
                                return null;
                            }
                            //Return the level
                            return DetermineHierarchyPropertyForQueryItems(workItemToUse, queryWorkItems, nameOfRelationship);
                        }
                        var fieldToEvaluate = workItemToUse.Fields[fieldName];
                        return fieldToEvaluate == null ? null : fieldToEvaluate.Value.ToString();
                    });

                if (config == null)
                {
                    SyncServiceTrace.D(Resources.ConfigurationNotExists);
                }
                if (_items != null && _items.ContainsKey(workItem.Id))
                {
                    config = _items[workItem.Id];
                }

                if (config != null)
                {
                    WorkItems.Add(new TfsWorkItem(workItem, this, fields, config));
                }
                _tfsItems.Add(workItem);
            }
        }

        /// <summary>
        ///Determine the Hierarchy Level of the current WorkItem in relation to the items of the current query
        /// </summary>
        /// <returns>A String, stating the level of the current Work item. e.G "0" for the root, "1" for a first level item...</returns>
        private string DetermineHierarchyPropertyForQueryItems(WorkItem workItemToUse, IList<WorkItem> queryWorkItems, string nameOfRelationship)
        {
            //Get the current parent
            var parentLink = workItemToUse.WorkItemLinks.Cast<WorkItemLink>().FirstOrDefault(x => x.LinkTypeEnd.Name.Equals(nameOfRelationship));
            //Check if there is a link to parent.
            //No parent --> First Level --> 0
            if (parentLink == null)
            {
                return "0";
            }
            var parentWorkItem = GetWorkItemFromQueryWorkItemsWithParentLink(queryWorkItems, parentLink);


            if (parentWorkItem != null)
            {
                //Check the Typ of the parent WorkItem. If the Type is different --> First Level --> 0
                if (parentWorkItem.Type != workItemToUse.Type)
                {
                    return "0";
                }
            }
            else
            {
                return "0";
            }
            // A parent Work Item exists
            // Iterate to the top of the tree and count the number of levels until there is no parent left or the parent is of a different type
            var currentLevel = 0;
            while ((parentWorkItem != null) && (workItemToUse.Type == parentWorkItem.Type))
            {
                parentLink = parentWorkItem.WorkItemLinks.Cast<WorkItemLink>().FirstOrDefault(x => x.LinkTypeEnd.Name == "Parent");
                parentWorkItem = GetWorkItemFromQueryWorkItemsWithParentLink(queryWorkItems, parentLink);
                ++currentLevel;
            }
            return currentLevel.ToString(CultureInfo.CurrentCulture);
        }

        /// <summary>
        /// Get a parent work item from a List using the parentLink
        /// The parent Work item is returned, or null if the link is null or no parent exists
        /// </summary>
        private static WorkItem GetWorkItemFromQueryWorkItemsWithParentLink(IList<WorkItem> queryWorkItems, WorkItemLink parentLink)
        {

            WorkItem parentWorkItem = null;
            //Check if the parent is in the current query
            if (parentLink != null)
            {
                foreach (WorkItem item in queryWorkItems)
                {
                    if (item.Id == parentLink.TargetId)
                    {
                        parentWorkItem = item;
                    }
                }
            }

            return parentWorkItem;
        }

        /// <summary>
        /// Loads all work items whose ids are given
        /// </summary>
        private IList<WorkItem> LoadWorkItemsWithoutConfiguration(int[] ids, IEnumerable<string> fields)
        {
            SyncServiceTrace.D(Resources.LoadWorkItemsWithoutConfiguration);
            var queryString = "SELECT System.Id FROM workitems";

            // When lazy loading missing work items for link format, it passes a list of field needed to
            // format the work item. Use these if passed, otherwise apply configuration
            var query = new Query(WorkItemStore, queryString);
            if (fields == null)
            {
                ApplyConfigurationFields(query);
            }
            else
            {
                query.DisplayFieldList.Clear();
                foreach (var field in fields)
                {
                    if (WorkItemStore.FieldDefinitions.Contains(field))
                    {
                        query.DisplayFieldList.Add(field);
                    }
                }
            }

            var result = ids == null
                             ? WorkItemStore.Query(new int[] { }, query.QueryString)
                             : WorkItemStore.Query(ids, query.QueryString);

            if (result != null && result.Count > 0)
            {
                SyncServiceTrace.D(Resources.NumberOfWorkItems + result.Count);
            }

            return (from WorkItem wi in result
                    select wi).ToList();
        }

        /// <summary>
        /// Loads all work items using a saved query
        /// </summary>
        public IList<WorkItem> LoadWorkItemsFromSavedQuery(IQueryConfiguration queryConfiguration)
        {
            SyncServiceTrace.D(Resources.LoadWorkItemsFromSavedQuery);
            var query = GetQuery(FindQueryInHierarchy(GetWorkItemQueries(), queryConfiguration.QueryPath));

            if (query == null)
            {
                var ex = new Exception(string.Format(Resources.Error_QueryDoesNotExists_Long, queryConfiguration.QueryPath));

                var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
                if (infoStorage != null)
                {
                    infoStorage.NotifyError(Resources.Error_QueryDoesNotExists_Short, ex);
                }

                throw ex;
            }

            ApplyConfigurationFields(query);

            if (query.IsLinkQuery)
            {
                // Load ids of linked work items and then
                // load them in a separate query
                var workItemLinkInfos = query.RunLinkQuery();
                var batchReadParams = new BatchReadParameterCollection();
                foreach (var linkInfo in workItemLinkInfos)
                {
                    if (!batchReadParams.Contains(linkInfo.TargetId))
                    {
                        batchReadParams.Add(new BatchReadParameter(linkInfo.TargetId));
                    }
                }

                var displayFieldPart = string.Join(",", query.DisplayFieldList.Cast<FieldDefinition>().Select(x => x.ReferenceName));
                var batchQuery = $"SELECT {displayFieldPart} FROM WorkItems";

                if (batchReadParams.Count != 0)
                {
                    var result = WorkItemStore.Query(batchReadParams, batchQuery);
                    return (from WorkItem wi in result select wi).ToList();
                }
            }
            else
            {
                var result = query.RunQuery();
                return (from WorkItem wi in result select wi).ToList();
            }

            return new List<WorkItem>();
        }

        /// <summary>
        /// Load all work items with one of the given ids
        /// </summary>
        private IList<WorkItem> LoadWorkItemsById(IQueryConfiguration queryConfig)
        {
            SyncServiceTrace.D(Resources.LoadWorkItemsById);
            var joinedIds = string.Join(",", queryConfig.ByIDs);
            var queryString = $"SELECT System.Id FROM workitems WHERE System.TeamProject='{ProjectName}' AND System.Id IN ({joinedIds})";

            var query = new Query(WorkItemStore, queryString);
            ApplyConfigurationFields(query);
            var workItemCollection = WorkItemStore.Query(query.QueryString);
            return (from WorkItem wi in workItemCollection
                    select wi).ToList();
        }

        /// <summary>
        /// Load all work items whose title contains a given text
        /// </summary>
        private IList<WorkItem> LoadWorkItemsByTitle(IQueryConfiguration queryConfig)
        {
            SyncServiceTrace.D(Resources.LoadWorkItemsByTitle);
            var queryString = $"SELECT System.Id FROM workitems WHERE System.TeamProject='{ProjectName}' AND System.Title='{queryConfig.ByTitle}'";
            if (string.IsNullOrEmpty(queryConfig.WorkItemType) == false)
            {
                queryString = $"SELECT System.Id FROM workitems WHERE System.TeamProject='{ProjectName}' AND System.WorkItemType='{queryConfig.WorkItemType}' AND System.Title='{queryConfig.ByTitle}'";
            }

            var query = new Query(WorkItemStore, queryString);
            ApplyConfigurationFields(query);
            var workItemCollection = WorkItemStore.Query(query.QueryString);
            return (from WorkItem wi in workItemCollection
                    select wi).ToList();
        }

        private void ApplyConfigurationFields(Query query)
        {
            query.DisplayFieldList.Clear();

            foreach (var fieldItem in _configuration.GetConfigurationItems()[0].FieldConfigurations)
            {
                if (WorkItemStore.FieldDefinitions.Contains(fieldItem.ReferenceFieldName))
                {
                    query.DisplayFieldList.Add(fieldItem.ReferenceFieldName);
                }
            }
        }

        /// <summary>
        /// Recursively searches for the query in the query hierarchy that has the given path
        /// </summary>
        private QueryDefinition FindQueryInHierarchy(IEnumerable<QueryItem> queryFolder, string queryPath)
        {
            if (queryFolder == null)
            {
                return null;
            }
            foreach (var item in queryFolder)
            {
                if (item is QueryDefinition && item.Path == queryPath)
                {
                    return (QueryDefinition)item;
                }

                if (item is QueryFolder)
                {
                    var foundItem = FindQueryInHierarchy(item as QueryFolder, queryPath);
                    if (foundItem != null)
                    {
                        return foundItem;
                    }
                }
            }

            return null;
        }

        private static IList<WorkItem> GetHierarchyWorkItemsByConfiguration(IList<WorkItem> workItems, IQueryConfiguration configuarion, WorkItemStore store)
        {
            // get link types from configuration
            IList<WorkItemLinkTypeEnd> configuartionLinkTypeEnds = null;
            if (configuarion.UseLinkedWorkItems)
            {
                configuartionLinkTypeEnds = new List<WorkItemLinkTypeEnd>();
                foreach (var linkType in configuarion.LinkTypes)
                {
                    var tfsLinkType = TfsWorkItemLinkType.GetTfsWorkItemLinkTypeByReferenceName(store, linkType);
                    if (tfsLinkType != null)
                    {
                        if (tfsLinkType.LinkType == Contracts.WorkItemObjects.LinkType.Full)
                        {
                            configuartionLinkTypeEnds.Add(tfsLinkType.WorkItemLinkType.ForwardEnd);
                            configuartionLinkTypeEnds.Add(tfsLinkType.WorkItemLinkType.ReverseEnd);
                        }
                        else if (tfsLinkType.LinkType == Contracts.WorkItemObjects.LinkType.ForwardEnd)
                        {
                            configuartionLinkTypeEnds.Add(tfsLinkType.WorkItemLinkType.ForwardEnd);
                        }
                        else if (tfsLinkType.LinkType == Contracts.WorkItemObjects.LinkType.ReverseEnd)
                        {
                            configuartionLinkTypeEnds.Add(tfsLinkType.WorkItemLinkType.ReverseEnd);
                        }
                    }
                }
            }

            var progressService = SyncServiceFactory.GetService<IProgressService>();
            progressService.EnterProgressGroup(workItems.Count, string.Format(CultureInfo.CurrentCulture, Resources.LoadWIs_Sorting, 1, workItems.Count));

            // go throu all query WorkItems
            IList<WorkItem> hierarchyWorkItems = new List<WorkItem>();
            while (workItems.Count > 0)
            {
                InsertLinkWorkItems(workItems, workItems[0], hierarchyWorkItems, configuartionLinkTypeEnds, progressService, configuarion);

                if (progressService.ProgressCanceled)
                {
                    return hierarchyWorkItems;
                }
            }

            return hierarchyWorkItems;
        }

        /// <summary>
        /// Sort and insert linked WorkItems.
        /// </summary>
        /// <param name="queryWorkItems">Raw, unsorted, possibly incomplete lisnameOfRelationshipt of work items.</param>
        /// <param name="parentWorkItem">Work item from queryWorkItems to use this iteration. The parent is removed from queryWorkItems and added to hierarchy work items, recursively followed by work items linked to it</param>
        /// <param name="hierarchyWorkItems">The generated sorted and complete list of work items.</param>
        /// <param name="configurationLinkTypeEnds">What link types to follow from parent work item and include in the list</param>
        /// <param name="progressService">progress service to publish progress to</param>
        private static void InsertLinkWorkItems(IList<WorkItem> queryWorkItems, WorkItem parentWorkItem,
                                                IList<WorkItem> hierarchyWorkItems, IList<WorkItemLinkTypeEnd> configurationLinkTypeEnds,
                                                IProgressService progressService, IQueryConfiguration configuration, bool isRecursion = false)
        {

            // if the sort array already contains WI then return
            if (progressService.ProgressCanceled || hierarchyWorkItems.FirstOrDefault(wi => wi.Id == parentWorkItem.Id) != null)
            {
                return;
            }

            progressService.DoTick(string.Format(CultureInfo.CurrentCulture, Resources.LoadWIs_Sorting, progressService.ActualProgressGroupActualTick, progressService.ActualProgressGroupCountOfTicks));

            // add new WI into sort array
            hierarchyWorkItems.Add(parentWorkItem);
            // remove WorkItem from query WorkItems collection
            queryWorkItems.Remove(parentWorkItem);

            // go through all links of current WI
            foreach (WorkItemLink workItemLink in parentWorkItem.WorkItemLinks)
            {
                if (progressService.ProgressCanceled)
                {
                    return;
                }

                WorkItem linkedWorkItem;
                if (configurationLinkTypeEnds == null)
                {
                    // if 'Include Linked Work Items' is checked off then search linked only in query WorkItems collection
                    linkedWorkItem = (from wi in queryWorkItems where wi.Id == workItemLink.TargetId select wi).FirstOrDefault();
                }
                else
                {
                    // if 'Include Linked Work Items' is checked on and are selected link types then check the link type
                    // against the selected link types in configuration
                    if (configurationLinkTypeEnds.Count > 0
                        && configurationLinkTypeEnds.Contains(workItemLink.LinkTypeEnd) == false)
                    {
                        continue;
                    }

                    // first try to find it in query WorkItems collection
                    linkedWorkItem = (from wi in queryWorkItems where wi.Id == workItemLink.TargetId select wi).FirstOrDefault();
                    // if not found then get WorkItem by ID
                    if (linkedWorkItem == null)
                    {
                        linkedWorkItem = parentWorkItem.Store.GetWorkItem(workItemLink.TargetId);
                    }
                }

                // if found then insert childs
                if (linkedWorkItem != null && !isRecursion)
                {
                    InsertLinkWorkItems(queryWorkItems, linkedWorkItem, hierarchyWorkItems, configurationLinkTypeEnds, progressService, configuration, configuration.IsDirectLinkOnlyMode);
                }
            }
        }

        #endregion Private methods
    }
}
