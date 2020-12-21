#region Usings
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using AIT.TFS.SyncService.Adapter.TFS2012;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Service.Configuration;
using AIT.TFS.SyncService.Service.Configuration.Serialization;
using Microsoft.TeamFoundation.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using TFS.Test.Common;
using Microsoft.TeamFoundation.Build.WebApi;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
#endregion

namespace TFS.SyncService.Test.Unit
{
    using AIT.TFS.SyncService.Adapter.TFS2012.BuildCenter;
    using AIT.TFS.SyncService.Contracts.BuildCenter;
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    ///This is a test class for the team foundation server adapter
    ///</summary>
    [TestClass]
    [DeploymentItem("Microsoft.WITDataStore32.dll")]
    [DeploymentItem("Microsoft.WITDataStore64.dll")]
    public class TfsSyncAdapterTest
    {
        #region Enums
        private enum BuildFilter { BuildNames, BuildAge, BuildQualities, BuildTags, All }
        #endregion

        #region Fields
        private Tfs2012SyncAdapter _tfsAdapter;
        private ServerConfiguration _serverConfiguration;
        private ServerConfiguration _serverConfigurationForBuildFilters;
        private Mock<IConfiguration> _configurationForBuildFilters;
        private Mock<IConfigurationTest> _configurationTestForBuildFilters;
        private Mock<IConfigurationTestResult> _configurationTestResultForBuildFilters;
        private BuildFilters _buildFilters;
        private static TestContext _testContext;        
        #endregion

        #region Properties

        /// <summary>
        /// Make sure all service are registered.
        /// </summary>
        [ClassInitialize]
        public static void MyClassInitialize(TestContext testContext)
        {
            _testContext = testContext;
            AIT.TFS.SyncService.Service.AssemblyInit.Instance.Init();
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);
        }

        /// <summary>
        /// Create an adapter connected to a 2012 server using standard configuration
        /// </summary>
        [TestInitialize]
        [TestCategory("ConnectionNeeded")]
        public void TestInitialize()
        {
            _serverConfiguration = CommonConfiguration.TfsTestServerConfiguration(_testContext);

            // ReSharper disable once UnusedVariable
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(_serverConfiguration.TeamProjectCollectionUrl));

            _tfsAdapter = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, _serverConfiguration.Configuration);

            _serverConfigurationForBuildFilters = new ServerConfiguration();
            _configurationForBuildFilters = new Mock<IConfiguration>();
            _configurationTestForBuildFilters = new Mock<IConfigurationTest>();
            _configurationTestResultForBuildFilters = new Mock<IConfigurationTestResult>();
            _buildFilters = new BuildFilters();
        }
        #endregion

        #region TestMethods

        /// <summary>
        ///A test for CreateNewWorkItem
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_CreateNewWorkItemTest_ShouldCreateNewWorkItem()
        {
            _tfsAdapter.Open(null);
            _tfsAdapter.CreateNewWorkItem(_serverConfiguration.Configuration.GetWorkItemConfiguration("Requirement"));

            Assert.AreEqual(1, _tfsAdapter.WorkItems.Count);
            Assert.AreEqual(_serverConfiguration.Configuration.GetWorkItemConfiguration("Requirement").FieldConfigurations.Count, _tfsAdapter.WorkItems.First().Fields.Count);
        }
   
        /// <summary>
        ///A test for GetAllWorkItemTypes
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_GetAllWorkItemTypes_ShouldReturnNonemptyList()
        {
            Assert.IsTrue(_tfsAdapter.GetAllWorkItemTypes().Count > 0);
        }

        /// <summary>
        ///A test for GetAreaPathHierarchy
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_GetAreaPathHierarchy_ShouldReturnNonemptyList()
        {
            // should have at least User Interface as child (for other tests)
            Assert.IsTrue(_tfsAdapter.GetAreaPathHierarchy().Childs.Count > 0);
        }

        /// <summary>
        /// Should have at least on other child (for other tests)
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_GetIterationPathHierarchy_ShouldReturnNonemptyList()
        {
            Assert.IsTrue(_tfsAdapter.GetIterationPathHierarchy().Childs.Count > 0);
        }

        /// <summary>
        ///A test for GetLinkTypes
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_GetLinkTypesTest_ShouldReturnNonemptyList()
        {
            Assert.IsTrue(_tfsAdapter.GetLinkTypes().Count > 0);
        }

        /// <summary>
        ///A test for GetWorkItemQueries
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_GetWorkItemQueriesTest_ShouldReturnNonemptyList()
        {
            Assert.IsTrue(_tfsAdapter.GetWorkItemQueries().Count > 0);
        }

        /// <summary>
        ///A test for Open
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_ShouldOpenItemsWithAdditionalFields()
        {
            _tfsAdapter.Open(new[]
                             {
                                 CommonConfiguration.RequirementId
                             },
                new[]
                {
                    "System.AssignedTo"
                });

            Assert.IsFalse(_serverConfiguration.Configuration.GetWorkItemConfiguration("Requirement").FieldConfigurations.Any(x => x.ReferenceFieldName == "System.AssignedTo"));
            Assert.IsTrue(_tfsAdapter.WorkItems.FirstOrDefault(x => x.Id == CommonConfiguration.RequirementId) != null);
            Assert.IsTrue(_tfsAdapter.WorkItems.First(x => x.Id == CommonConfiguration.RequirementId).Fields.Contains("System.AssignedTo"));
        }

        /// <summary>
        ///A test for Open
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_ShouldOpenWithStandardConfigurationTest()
        {
            Assert.IsTrue(_tfsAdapter.Open(new[]
                                           {
                                               CommonConfiguration.RequirementId, CommonConfiguration.IssueId
                                           }));
            Assert.AreEqual(2, _tfsAdapter.WorkItems.Count);
            Assert.IsTrue(_tfsAdapter.WorkItems.Find(CommonConfiguration.IssueId).Fields.Contains("Microsoft.VSTS.Common.Triage"));
        }

        /// <summary>
        /// Should use subtypes if configured.
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_ShouldUseSubtypeIfConfigured()
        {
            var requirementBaseConfiguration = new MappingElement
            {
                RelatedTemplate = "requirementBaseConfiguration",
                ImageFile = "standard.png",
                AssignRegularExpression = "Requirement",
                Converters = new MappingConverter[0],
                WorkItemType = "Requirement",
                MappingWorkItemType = "Requirement",
                AssignCellCol = 1,
                AssignCellRow = 1,
                Fields = new[]
                {
                    new MappingField
                    {
                        Direction = Direction.TfsToOther,
                        Name = "System.Id",
                        TableRow = 2,
                        TableCol = 3
                },
                    new MappingField
                    {
                        Direction = Direction.TfsToOther,
                        Name = "System.Rev",
                        TableRow = 2,
                        TableCol = 4
                    },
                    new MappingField
                    {
                        Direction = Direction.OtherToTfs,
                        Name = "System.Title",
                        TableRow = 2,
                        TableCol = 1
                    }
                }
            };

            var functionalRequirementConfiguration = new MappingElement
            {
                RelatedTemplate = "functionalRequirementConfiguration",
                ImageFile = "standard.png",
                AssignRegularExpression = "Requirement",
                Converters = new MappingConverter[0],
                WorkItemType = "Requirement",
                MappingWorkItemType = "Requirement",
                WorkItemSubtypeField = "Microsoft.VSTS.CMMI.RequirementType",
                WorkItemSubtypeValue = "Functional",
                AssignCellCol = 1,
                AssignCellRow = 1,
                Fields = new[]
                {
                    new MappingField
                    {
                        Direction = Direction.TfsToOther,
                        Name = "System.Id",
                        TableRow = 2,
                        TableCol = 3
                    },
                    new MappingField
                    {
                        Direction = Direction.TfsToOther,
                        Name = "System.Rev",
                        TableRow = 2,
                        TableCol = 4
                    },
                    new MappingField
                    {
                        Direction = Direction.OtherToTfs,
                        Name = "System.Title",
                        TableRow = 2,
                        TableCol = 1
                    }
                }
            };

            var configuration = new Configuration();
            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, requirementBaseConfiguration, string.Empty));
            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, functionalRequirementConfiguration, string.Empty));

            var target = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, configuration);
            Assert.IsTrue(target.Open(new[]
                                      {
                                          CommonConfiguration.RequirementId
                                      }));

            Assert.AreEqual("functionalRequirementConfiguration", target.WorkItems.Find(CommonConfiguration.RequirementId).Configuration.RelatedTemplate);
        }

        /// <summary>
        /// If the configuration provides a mapping that has no subtypes, this should be used in case the specific subtype is not configured.
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_ShouldDefaultToBaseConfigurationIfExactSubtypeIsNotFound()
        {
            var requirementBaseConfiguration = new MappingElement
            {
                RelatedTemplate = "requirementBaseConfiguration",
                ImageFile = "standard.png",
                AssignRegularExpression = "Requirement",
                Converters = new MappingConverter[0],
                WorkItemType = "Requirement",
                MappingWorkItemType = "Requirement",
                AssignCellCol = 1,
                AssignCellRow = 1,
                Fields = new[]
                {
                    new MappingField
                    {
                        Direction = Direction.TfsToOther,
                        Name = "System.Id",
                        TableRow = 2,
                        TableCol = 3
                    },
                    new MappingField
                    {
                        Direction = Direction.TfsToOther,
                        Name = "System.Rev",
                        TableRow = 2,
                        TableCol = 4
                    },
                    new MappingField
                    {
                        Direction = Direction.OtherToTfs,
                        Name = "System.Title",
                        TableRow = 2,
                        TableCol = 1
                    }
                }
            };

            var featureRequirementConfiguration = new MappingElement
            {
                RelatedTemplate = "featureRequirementConfiguration",
                ImageFile = "standard.png",
                AssignRegularExpression = "Requirement",
                Converters = new MappingConverter[0],
                WorkItemType = "Requirement",
                MappingWorkItemType = "Requirement",
                WorkItemSubtypeField = "Microsoft.VSTS.CMMI.RequirementType",
                WorkItemSubtypeValue = "Feature",
                AssignCellCol = 1,
                AssignCellRow = 1,
                Fields = new[]
                {
                    new MappingField
                    {
                        Direction = Direction.TfsToOther,
                        Name = "System.Id",
                        TableRow = 2,
                        TableCol = 3
                    },
                    new MappingField
                    {
                        Direction = Direction.TfsToOther,
                        Name = "System.Rev",
                        TableRow = 2,
                        TableCol = 4
                    },
                    new MappingField
                    {
                        Direction = Direction.OtherToTfs,
                        Name = "System.Title",
                        TableRow = 2,
                        TableCol = 1
                    }
                }
            };

            var configuration = new Configuration();
            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, requirementBaseConfiguration, string.Empty));
            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, featureRequirementConfiguration, string.Empty));

            var target = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, configuration);
            Assert.IsTrue(target.Open(new[]
            {
                CommonConfiguration.RequirementId
            }));

            Assert.AreEqual("requirementBaseConfiguration", target.WorkItems.Find(CommonConfiguration.RequirementId).Configuration.RelatedTemplate);
        }

        /// <summary>
        ///A test for OpenWithConfigurations
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_ShouldOpenWithConfigurations()
        {
            var config = new Configuration();
            var requirementConfig = _serverConfiguration.Configuration.GetWorkItemConfiguration("Requirement").Clone();
            requirementConfig.FieldConfigurations.Add(new ConfigurationFieldItem("Microsoft.VSTS.CMMI.RequirementType",
                "Microsoft.VSTS.CMMI.RequirementType",
                FieldValueType.PlainText,
                Direction.TfsToOther,
                0,
                0,
                string.Empty,
                false,
                HandleAsDocumentType.All,
                null,
                string.Empty,
                null,
                ShapeOnlyWorkaroundMode.AddSpace,
                false, null, null, null));
            config.GetConfigurationItems().Add(requirementConfig);

            var openConfig = new Dictionary<int, IConfigurationItem>
            {
                {
                    CommonConfiguration.RequirementId, requirementConfig
                }
            };

            var target = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, config);
            Assert.IsTrue(target.OpenWithConfigurations(openConfig));
            Assert.AreEqual(1, target.WorkItems.Count);
            Assert.IsTrue(target.WorkItems.Find(CommonConfiguration.RequirementId).Fields.Contains("Microsoft.VSTS.CMMI.RequirementType"));
        }

        /// <summary>
        ///A test for Save
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Save_ShouldSavePlaintext()
        {
            _tfsAdapter.Open(new[]
            {
                CommonConfiguration.RequirementId
            });
            var newTitle = "Title" + Guid.NewGuid();
            _tfsAdapter.WorkItems.First().Fields["System.Title"].Value = newTitle;

            Assert.AreEqual(0, _tfsAdapter.Save().Count);

            var target2 = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, _serverConfiguration.Configuration);
            target2.Open(new[]
            {
                CommonConfiguration.RequirementId
            });
            Assert.AreEqual(newTitle, target2.WorkItems.First().Title);
        }

        /// <summary>
        ///A test for Save
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_Save_SaveNewItemTest()
        {
            _tfsAdapter.Open(null);
            _tfsAdapter.CreateNewWorkItem(_serverConfiguration.Configuration.GetWorkItemConfiguration("Issue")).Fields["System.Title"].Value = "asdsdf";
            Assert.AreEqual(0, _tfsAdapter.Save().Count);
            Assert.AreNotEqual(0, _tfsAdapter.WorkItems.First().Id);

            var target2 = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, _serverConfiguration.Configuration);
            target2.Open(new[]
            {
                _tfsAdapter.WorkItems.First().Id
            });
            Assert.AreEqual(1, target2.WorkItems.Count);
        }

        /// <summary>
        ///A test for MappingFields
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_FieldDefinitions_ShouldReturnNonemptyList()
        {
            Assert.AreNotEqual(0, ((ITfsService)_tfsAdapter).FieldDefinitions);
        }

        /// <summary>
        ///A test for TestManagement
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_TestManagement_ShouldNotReturnNull()
        {
            Assert.IsNotNull(_tfsAdapter.TestManagement);
        }

        /// <summary>
        ///A test for WebAccess
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_GetWorkItemEditorUrl_UrlShouldBeValid()
        {
            // Arrange
            var computerName = _serverConfiguration.ComputerName;
            var fullyQualifiedComputerName = _serverConfiguration.FullyQualifiedComputerName;
            var workItemEditorUrl = _tfsAdapter.GetWorkItemEditorUrl(CommonConfiguration.RequirementId);
            var workItemEditorUrlFullyQualified = workItemEditorUrl.AbsoluteUri.Replace(computerName, fullyQualifiedComputerName);
            var request = WebRequest.Create(new Uri(workItemEditorUrlFullyQualified));
            request.PreAuthenticate = true;
            request.Credentials = CredentialCache.DefaultNetworkCredentials;

            // Act
            var response = (HttpWebResponse)request.GetResponse();

            // Assert
            Assert.AreEqual(HttpStatusCode.OK, response.StatusCode);
        }

        /// <summary>
        ///A test for RevisionWebAccessUrl
        ///</summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_RevisionWebAccessUrl_UrlShouldBeValid()
        {
            // Arrange
            var tfsService = (ITfsService)_tfsAdapter;
            var workItemMock = new Mock<IWorkItem>();
            workItemMock.SetupGet(x => x.Id).Returns(CommonConfiguration.TfsTestServerConfiguration(_testContext).HtmlRequirementId);
            var request = WebRequest.Create(tfsService.GetRevisionWebAccessUri(workItemMock.Object, 1));
            request.Credentials = CredentialCache.DefaultNetworkCredentials;

            // Act
            var response = (HttpWebResponse)request.GetResponse();

            // Assert
            Assert.AreEqual(HttpStatusCode.OK, response.StatusCode);
        }

        /// <summary>
        /// Checks, if the hierarchylevel field is set for the Work Items.
        /// Loads a single item --> Hierarchy Level 0
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_CheckHierarchyLevelFieldSingleItem()
        {
            const string hierarachyLevel0TemplateName = "HierarchieTemplateLevel0";
            const string hierarachyLevel1TemplateName = "HierarchieTemplateLevel1";

            //Create Configuration
            var featureBaseConfiguration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title").GetConfigurationItems().Single();
            var featureHierarchieLevel0Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel0TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "0").GetConfigurationItems().Single();
            var featureHierarchieLevel1Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel1TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "1").GetConfigurationItems().Single();

            //Add the configurations
            var configuration = new Configuration();
            configuration.TypeOfHierarchyRelationships = CommonConfiguration.RightNameOfHierarchyRelationship;
            configuration.GetConfigurationItems().Add(featureBaseConfiguration);
            configuration.GetConfigurationItems().Add(featureHierarchieLevel0Configuration);
            configuration.GetConfigurationItems().Add(featureHierarchieLevel1Configuration);

            var target = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, configuration);
            //Open the Element
            Assert.IsTrue(target.Open(new[] { CommonConfiguration.HierarchyTestChildId }));
            //Check it is not null
            Assert.IsNotNull(target.WorkItems.First());
            //Check if the right configuration has been selected
            Assert.AreEqual(hierarachyLevel0TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);
        }

        /// <summary>
        /// Checks, if the hierarchylevel field is set for the Work Items.
        /// Loads a two items
        /// First --> Hierarchy Level 0
        /// Second --> Hierarchy Level 1
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_CheckHierarchyLevelFieldMultipleItemsWithRightConfigurationOfHierarchyProperty()
        {
            const string hierarachyLevel0TemplateName = "HierarchieTemplateLevel0";
            const string hierarachyLevel1TemplateName = "HierarchieTemplateLevel1";

            //Create Configuration
            var featureBaseConfiguration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title").GetConfigurationItems().Single();
            var featureHierarchieLevel0Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel0TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "0").GetConfigurationItems().Single();
            var featureHierarchieLevel1Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel1TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "1").GetConfigurationItems().Single();

            //Add the configurations
            var configuration = new Configuration();

            configuration.TypeOfHierarchyRelationships = CommonConfiguration.RightNameOfHierarchyRelationship;

            configuration.GetConfigurationItems().Add(featureBaseConfiguration);
            configuration.GetConfigurationItems().Add(featureHierarchieLevel0Configuration);
            configuration.GetConfigurationItems().Add(featureHierarchieLevel1Configuration);

            var target = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, configuration);
            //Open the Element
            Assert.IsTrue(target.Open(new[] { CommonConfiguration.HierarchyTestParentId, CommonConfiguration.HierarchyTestChildId }));
            //Check it is not null
            Assert.IsNotNull(target.WorkItems.First());
            //Check the number of Work items
            Assert.IsTrue(target.WorkItems.Count == 2);

            //Check the links between the work items
            //Both Items should have the different templates, because the name of the relationship is right
            //Check if the right configuration has been selected work item 11 is the child of 10 --> HierarchyLevel 0
            Assert.AreEqual(hierarachyLevel0TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);

            //Check if the right configuration has been selected work item 10 is the parent of 11 --> HierarchyLevel1
            Assert.AreEqual(hierarachyLevel1TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);

            Assert.AreNotEqual("Generic10x1Table.xml", target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);
            Assert.AreNotEqual(hierarachyLevel1TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);

            Assert.AreNotEqual("Generic10x1Table.xml", target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);
            Assert.AreNotEqual(hierarachyLevel0TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);

        }



        /// <summary>
        /// Checks, if the hierarchylevel field is set for the Work Items.
        /// Loads a two items
        /// First --> Hierarchy Level 0
        /// Second --> Hierarchy Level 1
        /// Only the first HierarchyTemplates should be assigend, because the hierarchyproperty is set wrong
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_CheckHierarchyLevelFieldMultipleItemsWithWrongConfigurationOfHierarchyProperty()
        {
            const string hierarachyLevel0TemplateName = "HierarchieTemplateLevel0";
            const string hierarachyLevel1TemplateName = "HierarchieTemplateLevel1";

            //Create Configuration
            var featureBaseConfiguration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title").GetConfigurationItems().Single();
            var featureHierarchieLevel0Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel0TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "0").GetConfigurationItems().Single();
            var featureHierarchieLevel1Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel1TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "1").GetConfigurationItems().Single();

            //Add the configurations
            var configuration = new Configuration();

            configuration.TypeOfHierarchyRelationships = CommonConfiguration.WrongNameOfHierarchyRelationship;

            configuration.GetConfigurationItems().Add(featureBaseConfiguration);
            configuration.GetConfigurationItems().Add(featureHierarchieLevel0Configuration);
            configuration.GetConfigurationItems().Add(featureHierarchieLevel1Configuration);

            // if this test fails: make sure CommonConfiguration.RequirementId is of type "functional" !!
            var target = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, configuration);
            //Open the Element, This should work
            Assert.IsTrue(target.Open(new[] { CommonConfiguration.HierarchyTestParentId, CommonConfiguration.HierarchyTestChildId }));
            //Check it is not null
            Assert.IsNotNull(target.WorkItems.First());
            //Check the number of Work items
            Assert.IsTrue(target.WorkItems.Count == 2);

            //Check the links between the work items
            //Both Items should have the template of hierarchy level 0, because the name of the relationship is wrong
            //Check if the right configuration has been selected work item 11 is the child of 10
            Assert.AreEqual(hierarachyLevel0TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);

            //Check if the right configuration has been selected work item 10 is the parent of 11
            Assert.AreEqual(hierarachyLevel0TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);

            Assert.AreNotEqual("Generic10x1Table.xml", target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);
            Assert.AreNotEqual(hierarachyLevel1TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);

            Assert.AreNotEqual("Generic10x1Table.xml", target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);
            Assert.AreNotEqual(hierarachyLevel1TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);

        }

        /// <summary>
        /// Checks, if the hierarchylevel field is set for the Work Items.
        /// Loads a two items
        /// First --> Hierarchy Level 0
        /// Second --> Hierarchy Level 1
        /// HierarchyTemplates should not be assigend, because the hierarchy property is not set
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_CheckHierarchyLevelFieldMultipleItemsWithNoConfigurationOfHierarchyProperty()
        {
            const string hierarachyLevel0TemplateName = "HierarchieTemplateLevel0";
            const string hierarachyLevel1TemplateName = "HierarchieTemplateLevel1";

            //Create Configuration
            var featureBaseConfiguration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title").GetConfigurationItems().Single();
            var featureHierarchieLevel0Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel0TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "0").GetConfigurationItems().Single();
            var featureHierarchieLevel1Configuration = CommonConfiguration.GetWorkItemSubtypeConfiguration(hierarachyLevel1TemplateName, "Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title", "HierarchyLevel", "1").GetConfigurationItems().Single();

            //Add the configurations
            var configuration = new Configuration();

            configuration.TypeOfHierarchyRelationships = null;
            configuration.GetConfigurationItems().Add(featureBaseConfiguration);
            configuration.GetConfigurationItems().Add(featureHierarchieLevel0Configuration);
            configuration.GetConfigurationItems().Add(featureHierarchieLevel1Configuration);

            // if this test fails: make sure CommonConfiguration.RequirementId is of type "functional" !!
            var target = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, configuration);
            //Open the Element, This should work
            Assert.IsTrue(target.Open(new[] { CommonConfiguration.HierarchyTestParentId, CommonConfiguration.HierarchyTestChildId }));
            //Check it is not null
            Assert.IsNotNull(target.WorkItems.First());
            //Check the number of Work items
            Assert.IsTrue(target.WorkItems.Count == 2);

            //Check the links between the work items
            //Both Items should have the standard template, because there is no hierarchy property set
            //Check if the right configuration has been selected work item 11 is the child of 10
            Assert.AreEqual("Generic10x1Table.xml", target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);

            //Check if the right configuration has been selected work item 10 is the parent of 11
            Assert.AreEqual("Generic10x1Table.xml", target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);

            Assert.AreNotEqual(hierarachyLevel0TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);
            Assert.AreNotEqual(hierarachyLevel1TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);

            Assert.AreNotEqual(hierarachyLevel0TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);
            Assert.AreNotEqual(hierarachyLevel1TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);

        }

        /// <summary>
        /// Checks if the standard templates is assigend to the workitems if no hierarchietemplates are choosen
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_Open_CheckHierarchyLevelWithoutConfiguration()
        {
            const string hierarachyLevel0TemplateName = "HierarchieTemplateLevel0";
            const string hierarachyLevel1TemplateName = "HierarchieTemplateLevel1";

            //Create only the standard Configuration
            var featureBaseConfiguration = CommonConfiguration.GetSimpleFieldConfiguration("Requirement", Direction.OtherToTfs, FieldValueType.PlainText, "System.Title").GetConfigurationItems().Single();

            //Add the configurations
            var configuration = new Configuration();
            configuration.GetConfigurationItems().Add(featureBaseConfiguration);

            var target = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, configuration);
            //Open the Element
            Assert.IsTrue(target.Open(new[] { CommonConfiguration.HierarchyTestChildId, CommonConfiguration.HierarchyTestParentId }));
            //Check it is not null
            Assert.IsNotNull(target.WorkItems.First());
            //Check the number of Work items
            Assert.IsTrue(target.WorkItems.Count == 2);

            //Check the links between the work items

            //Check if the right configuration has been selected work item 11 is the child of 10
            Assert.AreEqual("Generic10x1Table.xml", target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);

            //Check if the right configuration has been selected work item 10 is the parent of 11
            Assert.AreEqual("Generic10x1Table.xml", target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);

            Assert.AreNotEqual(hierarachyLevel0TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);
            Assert.AreNotEqual(hierarachyLevel1TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestChildId).Configuration.RelatedTemplate);

            Assert.AreNotEqual(hierarachyLevel0TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);
            Assert.AreNotEqual(hierarachyLevel1TemplateName, target.WorkItems.Find(CommonConfiguration.HierarchyTestParentId).Configuration.RelatedTemplate);
        }

        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_TestResultExists_ShouldReturnFalse()
        {
            // Arange
            var tfsTestPlan = _tfsAdapter.AvailableTestPlans.First(f => f.Id == 1);

            var tfsTestSuite = tfsTestPlan.RootTestSuite.TestSuites.First(f => f.Id == 6);
            var tfsTestCase = _tfsAdapter.GetTestCases(tfsTestPlan, false).First(f => f.Id == 14);

            // Act
            var result = _tfsAdapter.TestResultExists(tfsTestPlan, tfsTestSuite, tfsTestCase, null, null);

            //Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_TestResultExists_ShouldReturnTrue()
        {
            // Arange
            var tfsTestPlan = _tfsAdapter.AvailableTestPlans.First(f => f.Id == 1);

            var tfsTestSuite = tfsTestPlan.RootTestSuite.TestSuites.First(f => f.Id == 6).TestSuites.First(f => f.Id == 8);
            var tfsTestCase = _tfsAdapter.GetTestCases(tfsTestPlan, false).First(f => f.Id == 14);

            // Act
            var result = _tfsAdapter.TestResultExists(tfsTestPlan, tfsTestSuite, tfsTestCase, null, null);

            //Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        [Ignore]
        // Todo: Implement Test
        public void TfsAdapter_SharedTestResultExists_ShouldReturnTrue()
        {
            // Similar to the two methods before, the shared test results functionality with TFS 2017 has to be tested
            // methods before: TfsAdapter_TestResultExists_ShouldReturnFalse() and TfsAdapter_TestResultExists_ShouldReturnTrue()
            // data to be tested: one test case in two different test suites, passed in one of them, marked as passed in both
        }

        [TestMethod]
        [TestCategory("ConnectionNeeded"), TestCategory("PreparationNeeded")]
        public void TfsAdapter_GetTestResults_ShouldCreateReportWithoutErrors()
        {
            //Arange
            var tfsTestPlan = _tfsAdapter.AvailableTestPlans.First(f => f.Id == 1);
            var tfsTestSuite = tfsTestPlan.RootTestSuite.TestSuites.First(f => f.Id == 6);
            var tfsTestCase = _tfsAdapter.GetTestCases(tfsTestPlan, false).First(f => f.Id == 29);
            Dictionary<int, int> lastTestRunPerConfig;

            //Act
            var result = _tfsAdapter.GetTestResults(tfsTestPlan, tfsTestCase.TestCase, null, null, tfsTestSuite, false, out lastTestRunPerConfig);

            //Assert
            Assert.IsNotNull(result);

        }

        /// <summary>
        /// Test data for testing method ApplyBuildFilters()
        /// </summary>
        /// <param name="buildFilter">Build filter</param>
        /// <returns>List of builds</returns>
        private List<ITfsServerBuild> testData(BuildFilter buildFilter)
        {

            var listOfBuilds = new List<ITfsServerBuild>();
            var firstBuild = new Build();
            firstBuild.BuildNumber = "HelloWorld_20170330.2";
            firstBuild.Quality = "Released";
            firstBuild.Definition = new DefinitionReference();
            firstBuild.Definition.Type = DefinitionType.Xaml;
            var dateTime = DateTime.Now;
            var olderDateTime = new DateTime(DateTime.Now.Year - 1, DateTime.Now.Month, DateTime.Now.Day);
            firstBuild.FinishTime = dateTime;

            var secondBuild = new Build();
            secondBuild.BuildNumber = "TestXaml";
            secondBuild.Quality = "Passed";
            secondBuild.Definition = new DefinitionReference();
            secondBuild.Definition.Type = DefinitionType.Xaml;
            secondBuild.FinishTime = dateTime;

            var thirdBuild = new Build();
            thirdBuild.BuildNumber = "TestBuild";
            thirdBuild.Definition = new DefinitionReference();
            thirdBuild.Definition.Type = DefinitionType.Build;
            thirdBuild.FinishTime = dateTime;

            switch (buildFilter)
            {
                case BuildFilter.BuildAge:
                    {
                        firstBuild.FinishTime = olderDateTime;
                        listOfBuilds.Add(new TfsServerBuild(secondBuild));
                        listOfBuilds.Add(new TfsServerBuild(firstBuild));
                        break;

                    }
                case BuildFilter.BuildNames:
                    {
                        listOfBuilds.Add(new TfsServerBuild(firstBuild));
                        listOfBuilds.Add(new TfsServerBuild(secondBuild));
                        break;
                    }
                case BuildFilter.BuildQualities:
                    {
                        listOfBuilds.Add(new TfsServerBuild(firstBuild));
                        listOfBuilds.Add(new TfsServerBuild(secondBuild));
                        listOfBuilds.Add(new TfsServerBuild(thirdBuild));
                        break;
                    }
                case BuildFilter.All:
                    {
                        secondBuild.FinishTime = olderDateTime;
                        thirdBuild.FinishTime = olderDateTime;
                        listOfBuilds.Add(new TfsServerBuild(firstBuild));
                        listOfBuilds.Add(new TfsServerBuild(secondBuild));
                        listOfBuilds.Add(new TfsServerBuild(thirdBuild));
                        break;
                    }
            }
            return listOfBuilds;
        }

        /// <summary>
        /// Checks if the method ApplyBuildFilters() returns correct list of builds acording to buildFilter( buildAge )
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_ApplyBuildFilters_ShouldReturnFilteredListOfBuildsAccordingToBuildAge()
        {
            //Arange
            var listOfBuilds = testData(BuildFilter.BuildAge);
            _buildFilters.BuildAge = "150";

            _configurationTestResultForBuildFilters.SetupGet(x => x.BuildFilters).Returns(_buildFilters);
            _configurationTestForBuildFilters.SetupGet(x => x.ConfigurationTestResult).Returns(_configurationTestResultForBuildFilters.Object);
            _configurationForBuildFilters.SetupGet(x => x.ConfigurationTest).Returns(_configurationTestForBuildFilters.Object);
            _serverConfigurationForBuildFilters.Configuration = _configurationForBuildFilters.Object;
            var tfsAdapter = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, _serverConfigurationForBuildFilters.Configuration);

            //Act
            var result = tfsAdapter.ApplyBuildFilters(listOfBuilds);

            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(result.First().BuildNumber, listOfBuilds.First().BuildNumber);
        }

        /// <summary>
        /// Checks if the method ApplyBuildFilters() returns correct list of builds acording to buildFilter( buildNames )
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        public void TfsAdapter_ApplyBuildFilters_ShouldReturnFilteredListOfBuildsAccordingToBuildNames()
        {
            //Arange
            var listOfBuilds = testData(BuildFilter.BuildNames);
            var listOfBuildNames = new List<string>();
            listOfBuildNames.Add("HelloWorld_20170330.2");
            _buildFilters.BuildNames = listOfBuildNames;

            _configurationTestResultForBuildFilters.SetupGet(x => x.BuildFilters).Returns(_buildFilters);
            _configurationTestForBuildFilters.SetupGet(x => x.ConfigurationTestResult).Returns(_configurationTestResultForBuildFilters.Object);
            _configurationForBuildFilters.SetupGet(x => x.ConfigurationTest).Returns(_configurationTestForBuildFilters.Object);
            _serverConfigurationForBuildFilters.Configuration = _configurationForBuildFilters.Object;
            var tfsAdapter = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, _serverConfigurationForBuildFilters.Configuration);

            //Act
            var result = tfsAdapter.ApplyBuildFilters(listOfBuilds);

            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(result.First().BuildNumber, listOfBuilds.First().BuildNumber);
        }

        /// <summary>
        /// Checks if the method ApplyBuildFilters() returns correct list of builds acording to buildFilter( buildQuailties )
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public void TfsAdapter_ApplyBuildFilters_ShouldReturnFilteredListOfBuildsAccordingToBuildQualities()
        {
            //Arange
            var listOfBuilds = testData(BuildFilter.BuildQualities);

            var listOfBuildQualities = new List<string>();
            listOfBuildQualities.Add("Released");
            _buildFilters.BuildQualities = listOfBuildQualities;

            _configurationTestResultForBuildFilters.SetupGet(x => x.BuildFilters).Returns(_buildFilters);
            _configurationTestForBuildFilters.SetupGet(x => x.ConfigurationTestResult).Returns(_configurationTestResultForBuildFilters.Object);
            _configurationForBuildFilters.SetupGet(x => x.ConfigurationTest).Returns(_configurationTestForBuildFilters.Object);
            _serverConfigurationForBuildFilters.Configuration = _configurationForBuildFilters.Object;
            var tfsAdapter = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, _serverConfigurationForBuildFilters.Configuration);

            //Act
            var result = tfsAdapter.ApplyBuildFilters(listOfBuilds);

            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(result.FirstOrDefault().BuildNumber, listOfBuilds.FirstOrDefault().BuildNumber);
            Assert.IsTrue(result.Count == 2);
        }

        /// <summary>
        /// Checks if the method ApplyBuildFilters() returns correct list of builds acording to buildFilters ( all filters )
        /// </summary>
        [TestMethod]
        [TestCategory("ConnectionNeeded")]
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public void TfsAdapter_ApplyBuildFilters_ShouldReturnFilteredListOfBuildsAccordingToAllBuildFilters()
        {
            //Arange
            var listOfBuilds = testData(BuildFilter.All);

            var listOfBuildQualities = new List<string>();
            listOfBuildQualities.Add("Released");
            listOfBuildQualities.Add("Passed");
            _buildFilters.BuildQualities = listOfBuildQualities;

            var listOfBuildNames = new List<string>();
            listOfBuildNames.Add("HelloWorld_20170330.2");
            listOfBuildNames.Add("TestBuild");
            _buildFilters.BuildNames = listOfBuildNames;

            _buildFilters.BuildAge = "150";

            _configurationTestResultForBuildFilters.SetupGet(x => x.BuildFilters).Returns(_buildFilters);
            _configurationTestForBuildFilters.SetupGet(x => x.ConfigurationTestResult).Returns(_configurationTestResultForBuildFilters.Object);
            _configurationForBuildFilters.SetupGet(x => x.ConfigurationTest).Returns(_configurationTestForBuildFilters.Object);
            _serverConfigurationForBuildFilters.Configuration = _configurationForBuildFilters.Object;
            var tfsAdapter = new Tfs2012SyncAdapter(_serverConfiguration.TeamProjectCollectionUrl, _serverConfiguration.TeamProjectName, null, _serverConfigurationForBuildFilters.Configuration);

            //Act
            var result = tfsAdapter.ApplyBuildFilters(listOfBuilds);

            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(result.FirstOrDefault().BuildNumber, listOfBuilds.FirstOrDefault().BuildNumber);
            Assert.IsTrue(result.Count == 1);
        }
        #endregion

    }
}

