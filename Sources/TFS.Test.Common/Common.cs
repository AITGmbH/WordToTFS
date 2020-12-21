#region Usings
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Service.Configuration;
using AIT.TFS.SyncService.Service.Configuration.Serialization;
using Moq;
#endregion

namespace TFS.Test.Common
{
    using System.Diagnostics;
    using Microsoft.VisualStudio.Services.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Common data for all unit tests
    /// Server data and configuration
    /// </summary>
    public static class CommonConfiguration
    {
        //private static MappingConverter converter = new MappingConverter();
        // TODO Extend unit test to converters

        private static readonly string _dllLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        /// <summary>
        /// Requirement work item type configuration
        /// </summary>
        private static MappingElement GetRequirement(string htmlFieldReferenceName)
        {
            return new MappingElement
            {
                RelatedTemplate = "Configuration\\Requirement.xml",
                ImageFile = "standard.png",
                AssignRegularExpression = "Requirement",
                Converters = new MappingConverter[0],
                WorkItemType = "Requirement",
                MappingWorkItemType = "Requirement",
                AssignCellCol = 1,
                AssignCellRow = 1,

                Fields = new[]
                {
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.Id", TableRow = 2, TableCol = 3, DefaultValue = new MappingFieldDefaultValue { DefaultValue = "ABCD"}},
                    new MappingField { Direction = Direction.TfsToOther, Name ="System.Rev", TableRow = 2, TableCol = 4},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.Title", TableRow = 2, TableCol = 1},
                    new MappingField { Direction = Direction.OtherToTfs, FieldValueType = FieldValueType.HTML, Name=htmlFieldReferenceName, TableRow = 3, TableCol = 1, WordBookmark = BookmarkName},
                    new MappingField { Direction = Direction.OtherToTfs, Name="Microsoft.VSTS.Common.Triage", TableRow = 7, TableCol = 1},
                    new MappingField { Direction = Direction.OtherToTfs, FieldValueType = FieldValueType.DropDownList, Name="System.State", TableRow = 1, TableCol = 4},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.AreaPath", TableRow = 4, TableCol = 1},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.IterationPath", TableRow = 5, TableCol = 1}
                },

                Links = new[]
                {
                    new MappingLink { Direction = Direction.GetOnly, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 1},
                    new MappingLink { Direction = Direction.OtherToTfs, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 2},
                    new MappingLink { Direction = Direction.PublishOnly, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 3},
                    new MappingLink { Direction = Direction.TfsToOther, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 4},
                    new MappingLink { Direction = Direction.SetInNewTfsWorkItem, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 5},

                    new MappingLink { Direction = Direction.OtherToTfs, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Child", Overwrite = true, TableRow = 2, TableCol = 2}
                }
            };
        }


        /// <summary>
        /// Requirement work item type configuration
        /// </summary>
        private static MappingElement GetRequirementwithVariable(string htmlFieldReferenceName)
        {
            return new MappingElement
            {
                RelatedTemplate = "Configuration\\RequirementWithVariable.xml",
                //RelatedTemplate = _dllLocation + "\\Configuration\\RequirementWithVariable.xml",

                ImageFile = "standard.png",
                AssignRegularExpression = "Requirement",
                Converters = new MappingConverter[0],
                WorkItemType = "Requirement",
                MappingWorkItemType = "Requirement",
                AssignCellCol = 1,
                AssignCellRow = 1,

                Fields = new[]
                {
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.Id", TableRow = 2, TableCol = 3, DefaultValue = new MappingFieldDefaultValue { DefaultValue = "ABCD"}},
                    new MappingField { Direction = Direction.TfsToOther, Name ="System.Rev", TableRow = 2, TableCol = 4},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.Title", TableRow = 2, TableCol = 1},
                    new MappingField { Direction = Direction.OtherToTfs, FieldValueType = FieldValueType.HTML, Name=htmlFieldReferenceName, TableRow = 3, TableCol = 1, WordBookmark = BookmarkName},
                    new MappingField { Direction = Direction.OtherToTfs, Name="Microsoft.VSTS.Common.Triage", TableRow = 7, TableCol = 1},
                    new MappingField { Direction = Direction.OtherToTfs, FieldValueType = FieldValueType.DropDownList, Name="System.State", TableRow = 1, TableCol = 4},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.AreaPath", TableRow = 4, TableCol = 1},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.IterationPath", TableRow = 5, TableCol = 1},
                    new MappingField { Direction = Direction.OtherToTfs, Name="VariableTest", VariableName = TestVariableName, TableRow = 8, TableCol = 1,  FieldValueType = FieldValueType.BasedOnVariable, WordBookmark = StaticValueTextBookMarkName}
                },

                Links = new[]
                {
                    new MappingLink { Direction = Direction.GetOnly, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 1},
                    new MappingLink { Direction = Direction.OtherToTfs, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 2},
                    new MappingLink { Direction = Direction.PublishOnly, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 3},
                    new MappingLink { Direction = Direction.TfsToOther, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 4},
                    new MappingLink { Direction = Direction.SetInNewTfsWorkItem, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Affects", Overwrite = false, TableRow = 6, TableCol = 5},

                    new MappingLink { Direction = Direction.OtherToTfs, LinkFormat = "{System.Id}", LinkSeparator = ",", LinkValueType = "Child", Overwrite = true, TableRow = 2, TableCol = 2}
                }
            };
        }

        /// <summary>
        /// Requirement work item type configuration modeled on CMMI (2015) template
        /// </summary>
        public static IConfiguration GetCMMIRequirement()
        {
            var configuration = new Configuration();

            var item = new MappingElement

            {
                RelatedTemplate = "Configuration\\Requirement.xml",
                ImageFile = "standard.png",
                AssignRegularExpression = "Requirement",
                Converters = new MappingConverter[0],
                WorkItemType = "Requirement",
                MappingWorkItemType = "Requirement",
                AssignCellCol = 1,
                AssignCellRow = 1,

                Fields = new[]
                {
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.Title", TableRow = 1, TableCol = 2},
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.Id", TableRow = 1, TableCol = 3},
                    new MappingField { Direction = Direction.TfsToOther, Name ="System.Rev", TableRow = 1, TableCol = 4},
                    new MappingField { Direction = Direction.OtherToTfs, Name="Microsoft.VSTS.Scheduling.Size", TableRow = 2, TableCol = 1},
                    new MappingField { Direction = Direction.OtherToTfs, Name="Microsoft.VSTS.CMMI.RequirementType", TableRow = 2, TableCol = 2, FieldValueType = FieldValueType.DropDownList},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.AssignedTo", FieldValueType = FieldValueType.DropDownList, TableRow = 2, TableCol = 3},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.State", FieldValueType = FieldValueType.DropDownList, TableRow = 2, TableCol = 4},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.Description", TableRow = 3, TableCol = 1, HandleAsDocument = true, HandleAsDocumentMode = HandleAsDocumentType.OleOnDemand, OLEMarkerField = null, OLEMarkerValue = string.Empty, ShapeOnlyWorkaroundMode = ShapeOnlyWorkaroundMode.AddSpace}
                }
            };

            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, item, Path.Combine(Environment.CurrentDirectory, "Configuration")));
            return configuration;
        }


        /// <summary>
        /// Issue work item type configuration.
        /// </summary>
        private static MappingElement Issue = new MappingElement
        {

            RelatedTemplate = "Configuration\\Issue.xml",
            //RelatedTemplate = _dllLocation + "\\Configuration\\Issue.xml",

            ImageFile = "standard.png",
            Converters = new MappingConverter[0],
            WorkItemType = "Issue",
            MappingWorkItemType = "Issue",
            AssignRegularExpression = "Issue",
            AssignCellCol = 1,
            AssignCellRow = 1,
            Fields = new[]
            {
                new MappingField { Direction = Direction.TfsToOther, Name = "System.Id"},
                new MappingField { Direction = Direction.TfsToOther, Name ="System.Rev"},
                new MappingField { Direction = Direction.OtherToTfs, Name="System.Title"},
                new MappingField { Direction = Direction.OtherToTfs, Name="Microsoft.VSTS.Common.Triage"},
                new MappingField { Direction = Direction.OtherToTfs, Name="System.Description", HandleAsDocument = true, HandleAsDocumentMode = HandleAsDocumentType.All, OLEMarkerField = null, OLEMarkerValue = string.Empty}
            }
        };

        /// <summary>
        /// Test case work item type configuration.
        /// </summary>
        private static MappingElement TestCase = new MappingElement
        {

            RelatedTemplate = "Configuration\\TestCase.xml",

            ImageFile = "standard.png",
            Converters = new MappingConverter[0],
            WorkItemType = "Test Case",
            MappingWorkItemType = "Test Case",
            AssignRegularExpression = "Test Case",
            AssignCellCol = 1,
            AssignCellRow = 1,

            Fields = new[]
            {
                new MappingField { Direction = Direction.TfsToOther, Name = "System.Id", TableRow = 2, TableCol = 3, DefaultValue = new MappingFieldDefaultValue { DefaultValue = "ABCD"}},
                new MappingField { Direction = Direction.TfsToOther, Name ="System.Rev", TableRow = 2, TableCol = 4},
                new MappingField { Direction = Direction.OtherToTfs, Name="System.Title", TableRow = 2, TableCol = 1},
                new MappingField { Direction = Direction.OtherToTfs, FieldValueType= FieldValueType.BasedOnFieldType, Name="Microsoft.VSTS.TCM.Steps", TestCaseStepDelimiter="->", TableRow = 3, TableCol = 1}
            }
        };

        /// <summary>
        /// A first level header that maps System.AreaPath, System.IterationPath and System.Title
        /// </summary>
        public static MappingHeader Header1 = new MappingHeader
        {
            RelatedTemplate = string.Empty,
            ImageFile = "standard.png",
            Converters = new MappingConverter[0],
            Identifier = "1Header",
            Level = 1,
            Column = 1,
            Row = 1,
            Fields = new[]
            {
                new MappingField { Direction = Direction.OtherToTfs, Name = "System.AreaPath", TableCol=1, TableRow=2},
                new MappingField { Direction = Direction.OtherToTfs, Name ="System.IterationPath", TableCol = 1, TableRow=3},
                new MappingField { Direction = Direction.SetInNewTfsWorkItem, Name ="System.Title", TableCol = 1, TableRow=4}
            }
        };

        /// <summary>
        /// A first level header without field mappings. Its purpose is only to stop other first level headers.
        /// </summary>
        private static MappingHeader Header1End = new MappingHeader
        {
            RelatedTemplate = string.Empty,
            ImageFile = "standard.png",
            Converters = new MappingConverter[0],
            Identifier = "1HeaderEnd",
            Level = 1,
            Column = 1,
            Row = 1,
            Fields = new MappingField[0]
        };

        /// <summary>
        /// A second level header that also maps a cell to a field that does not exist.
        /// </summary>
        public static MappingHeader Header2 = new MappingHeader
        {
            RelatedTemplate = string.Empty,
            ImageFile = "standard.png",
            Converters = new MappingConverter[0],
            Identifier = "2Header",
            Level = 2,
            Column = 1,
            Row = 1,
            Fields = new[]
            {
                new MappingField { Direction = Direction.SetInNewTfsWorkItem, Name = "System.AreaPath", TableCol = 1, TableRow = 2},
                new MappingField { Direction = Direction.OtherToTfs, Name = "System.GibtsNicht", TableCol = 1, TableRow=3}
            }
        };

        /// <summary>
        /// Configuration used for testing purposes.
        /// </summary>
        public static IConfiguration Configuration
        {
            get
            {
                return GetConfiguration("System.Description", false);
            }
        }

        /// <summary>
        /// Configuration used for testing purposes.
        /// Returns a configuration that contains a variables that is used by a static value field
        /// </summary>
        public static IConfiguration ConfigurationWithVariables
        {
            get
            {
                return GetConfiguration("System.Description", true);
            }
        }

        /// <summary>
        /// Configuration used for testing purposes.
        /// </summary>
        public static IConfiguration GetConfiguration(string htmlFieldReferenceName, bool getRequirementWithVariable)
        {
            var configuration = new Configuration();

            if (getRequirementWithVariable)
            {
                //Add the variable to the configuartion
                var variable = new Variable();
                variable.Name = TestVariableName;
                variable.Value = TestVariableText;
                configuration.Variables.Add(variable);
                configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, GetRequirementwithVariable(htmlFieldReferenceName), string.Empty));
            }
            else
            {
                configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, GetRequirement(htmlFieldReferenceName), string.Empty));
            }

            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, Issue, Environment.CurrentDirectory));
            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, TestCase, string.Empty));

            configuration.Headers.Add(new ConfigurationItem(configuration, Header1, string.Empty));
            configuration.Headers.Add(new ConfigurationItem(configuration, Header2, string.Empty));
            configuration.Headers.Add(new ConfigurationItem(configuration, Header1End, string.Empty));

            return configuration;
        }

        /// <summary>
        /// Gets a mock configuration.
        /// </summary>
        public static Mock<IConfiguration> MockConfiguration
        {
            get
            {
                var config = new Mock<IConfiguration>
                {
                    DefaultValue = DefaultValue.Mock
                };
                config.SetupGet(x => x.DefaultMapping).Returns(string.Empty);
                config.SetupGet(x => x.MappingFolder).Returns(string.Empty);
                var configurationItems = new List<IConfigurationItem>
                                             {
                                                 new ConfigurationItem(config.Object, GetRequirement("System.Description"), string.Empty),
                                                 new ConfigurationItem(config.Object, Issue, string.Empty),
                                                 new ConfigurationItem(config.Object, Header1End, string.Empty)
                                             };

                var headerItems = new List<IConfigurationItem>
                                             {
                                                 new ConfigurationItem(config.Object, Header1, string.Empty),
                                                 new ConfigurationItem(config.Object, Header2, string.Empty),
                                                 new ConfigurationItem(config.Object, TestCase, string.Empty)
                                             };

                config.Setup(x => x.GetConfigurationItems()).Returns(configurationItems);
                config.SetupGet(x => x.Headers).Returns(headerItems);
                return config;
            }
        }

        /// <summary>
        /// Id of a requirement work item type used for testing purposes.
        /// The requirement must have an "affects" link to the "RequirementsAffectedItemId".
        /// The requirement must have "Type" = 'Functional'
        /// </summary>
        public const int RequirementId = 22;

        /// <summary>
        /// Id of an issue work item type. Any issue will do.
        /// </summary>
        public const int IssueId = 24;

        /// <summary>
        /// Id of a work item that is not located in the Team Project for these unit tests.
        /// </summary>
        public const int IdOfWorkItemInOtherTeamProject = 3;

        /// <summary>
        /// Id of a work item that the "RequirementId" links to via an "affects" link
        /// </summary>
        public const int RequirementAffectedItemId = 23;

        /// <summary>
        /// Id of a work item used to test links. Any work item different from the others will do
        /// </summary>
        public const int LinkTestWorkItemId = 24;

        /// <summary>
        /// Id of a test case work item type. Any test case will do. This test case will have it steps modified
        /// </summary>
        public const int TestCaseId = 15;

        /// <summary>
        /// Id of a test suite.
        /// </summary>
        public const int TestSuiteId = 6;

        /// <summary>
        /// Id of a test plan. 
        /// </summary>
        public const int TestPlanId = 1;

        /// <summary>
        /// Id of a Test Case work item type. This Test Case has some Build information attached to it
        /// </summary>
        public const int TestCaseIdWithBuildInformation = 237;

        /// <summary>
        /// Id of a Test Case work item type. This Test Case has a Test Result within a Test Run in which a Bug was created.
        /// </summary>
        public const int TestCaseIdWithTestResultLinkedToBug = 73;

        /// <summary>
        /// Id of a Test Suite. This testsuite contains a Test Case  with build information
        /// </summary>
        public const int TestSuiteIdWithBuildInformation = 20;

        /// <summary>
        /// Id of a Test Plan. This test plan contains a Test Case with build information
        /// </summary>
        public const int TestPlanIdWithBuildInformation = 14;
        /// <summary>
        /// A working name of a hierarchyrelationship
        /// </summary>
        public const string RightNameOfHierarchyRelationship = "Parent";

        /// <summary>
        /// A non working name of a hierarchyrelationship
        /// </summary>
        public const string WrongNameOfHierarchyRelationship = "WrongParent";

        /// <summary>
        /// Id of a Requirement that acts as a parent node
        /// </summary>
        public const int HierarchyTestParentId = 27;

        /// <summary>
        /// Id of a Requirement that acts as a child node
        /// </summary>
        public const int HierarchyTestChildId = 28;

        /// <summary>
        /// Name of the Bookmark
        /// </summary>
        public const string BookmarkName = "NameOfARandomBookMark";

        /// <summary>
        /// Name of the Bookmark that is used to test static value field
        /// </summary>
        public const string StaticValueTextBookMarkName = "NameOfTheBookMarkToTestStaticValueFields";

        public const string TestVariableName = "SampleVariable";

        public const string TestVariableText = "SomeTextForTheSampleVariable";


        /// <summary>
        /// Ids on the new test structure
        /// </summary>
        public const int TestPlan_1_Id = 1;
        public const int TestSuite_1_1_TestSuiteId = 6;
        public const int TestSuite_1_1_1_TestSuiteId = 8;
        public const int TestSuite_1_2_TestSuiteId = 7;
        public const int TestSuite_1_2_1TestSuiteId = 9;

        public const int TestPlan_2_Id = 3;
        public const int TestSuite_2_1_TestSuiteId = 10;
        public const int TestSuite_2_1_1_TestSuiteId = 11;
        public const int TestSuite_2_2_TestSuiteId = 12;
        public const int TestSuite_2_2_1TestSuiteId = 13;

        public const int TestCase_InAllSuitesAndPlans_Id = 14;
        public const int TestCase_UnderTestPlan1_Id = 5;
        public const int TestCase_1_1_Id = 15;
        public const int TestCase_1_1_1_Id = 17;
        public const int TestCase_1_2_Id = 18;
        public const int TestCase_1_2_1_Id = 19;

        public const int ConfigurationId_ThatExistsInAllTestCases = 2;
        public const int ConfigurationName_1 = 7;
        public const int ConfigurationName_1_1 = 3;
        public const int ConfigurationName_1_1_1 = 5;
        public const int ConfigurationName_1_2 = 4;
        public const int ConfigurationName_1_2_1 = 6;

        /// <summary>
        /// Configuration data for TFS2017
        /// </summary>
        public static ServerConfiguration TfsTestServerConfiguration(TestContext testContext)
        {
            var serverConfig = new ServerConfiguration
            {
                ComputerName = testContext.Properties["TfsServerName"].ToString(),
                FullyQualifiedComputerName = testContext.Properties["TfsServerFQDN"].ToString(),
                TeamProjectCollectionUrl = testContext.Properties["TfsTeamProjectCollectionUrl"].ToString(),
                TeamProjectCollectionName = testContext.Properties["TfsTeamProjectCollectionName"].ToString(),
                TeamProjectName = testContext.Properties["TeamProjectName"].ToString(),
                HtmlFieldReferenceName = "System.Description",
                Configuration = GetConfiguration("System.Description", false),
                HtmlRequirementId = 26, // id of WI with HTML System.Description
                SynchronizedWorkItemId = 83
            };
            return serverConfig;
        }

        /// <summary>
        /// Replacing configuration tokens in W2T config files
        /// </summary>
        /// <param name="testcontext"></param>
        public static void ReplaceConfigFileTokens(TestContext testcontext)
        {

            IList<string> configFiles = new List<string>();

            try
            {
                configFiles.AddRange(Directory.GetFiles(_dllLocation + "\\Configuration", "Config*.xml"));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            try
            {
                configFiles.AddRange(Directory.GetFiles(_dllLocation + "\\Config", "Config*.xml"));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }


            foreach (var configFile in configFiles)
            {
                var content = File.ReadAllText(configFile);
                var teamProjectCollectionUrlPlaceholder = testcontext.Properties["TeamProjectCollectionUrlPlaceholder"].ToString();
                if (content.Contains(teamProjectCollectionUrlPlaceholder))
                {
                    content = content.Replace(teamProjectCollectionUrlPlaceholder, testcontext.Properties["TfsTeamProjectCollectionUrl"].ToString());
                    File.WriteAllText(configFile, content);
                }

                var teamProjectNamePlaceholder = testcontext.Properties["TeamProjectNamePlaceholder"].ToString();
                if (content.Contains(teamProjectNamePlaceholder))
                {
                    content = content.Replace(teamProjectNamePlaceholder, testcontext.Properties["TeamProjectName"].ToString());
                    File.WriteAllText(configFile, content);
                }
            }
        }

        /// <summary>
        /// Create a configuration with exactly one work item that has ID, Rev and a given field.
        /// </summary>
        /// <param name="workItemType">Configuration work item type</param>
        /// <param name="direction">Field of the custom field of the single work item.</param>
        /// <param name="fieldValueType">Value type of the custom field of the single work item.</param>
        /// <param name="fieldName">Field reference name of the field of the single work item.</param>
        /// <returns>Configuration containing exactly one work item configuration.</returns>
        public static IConfiguration GetSimpleFieldConfiguration(string workItemType, Direction direction, FieldValueType fieldValueType, string fieldName)
        {
            var configuration = new Configuration();
            var item = new MappingElement
            {
                RelatedTemplate = "Generic10x1Table.xml",
                ImageFile = "standard.png",
                AssignRegularExpression = workItemType,
                Converters = new MappingConverter[0],
                WorkItemType = workItemType,
                MappingWorkItemType = workItemType,
                AssignCellCol = 1,
                AssignCellRow = 1,
                Fields = new[]
                {
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.ID", TableRow = 2, TableCol = 1 },
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.Rev", TableRow = 3, TableCol = 1 },
                    new MappingField { Direction = direction, FieldValueType = fieldValueType, Name = fieldName, TableRow = 4, TableCol = 1 }

                },

                Links = new MappingLink[0]
            };

            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, item, Path.Combine(Environment.CurrentDirectory, "Configuration")));
            return configuration;
        }

        /// <summary>
        /// Create a configuration with exactly one work item that has ID, Rev and a given field.
        /// </summary>
        /// <param name="workItemType">Configuration work item type</param>
        /// <param name="direction">Field of the custom field of the single work item.</param>
        /// <param name="fieldValueType">Value type of the custom field of the single work item.</param>
        /// <param name="fieldName">Field reference name of the field of the single work item.</param>
        /// <returns>Configuration containing exactly one work item configuration.</returns>
        public static IConfiguration GetSimpleFieldConfigurationWithVariable(string workItemType, Direction direction, FieldValueType fieldValueType, string fieldName)
        {
            var configuration = new Configuration();

            //Add the variable to the configuartion
            var variable = new Variable();
            variable.Name = TestVariableName;
            variable.Value = TestVariableText;
            configuration.Variables.Add(variable);

            var item = new MappingElement
            {
                RelatedTemplate = "Generic10x1Table.xml",
                ImageFile = "standard.png",
                AssignRegularExpression = workItemType,
                Converters = new MappingConverter[0],
                WorkItemType = workItemType,
                MappingWorkItemType = workItemType,
                AssignCellCol = 1,
                AssignCellRow = 1,
                Fields = new[]
                {
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.ID", TableRow = 2, TableCol = 1 },
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.Rev", TableRow = 3, TableCol = 1 },
                    new MappingField { Direction = direction, FieldValueType = fieldValueType, Name = fieldName, TableRow = 4, TableCol = 1 },
                    new MappingField { Direction = Direction.OtherToTfs, Name="VariableTest", VariableName = TestVariableName, TableRow = 5, TableCol = 1,  FieldValueType = FieldValueType.BasedOnVariable, WordBookmark = StaticValueTextBookMarkName}

                },

                Links = new MappingLink[0]
            };

            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, item, Path.Combine(Environment.CurrentDirectory, "Configuration")));
            return configuration;
        }


        /// <summary>
        /// Create a configuration with exactly one work item that has ID, Rev and a given field.
        /// </summary>
        /// <param name="workItemType">Configuration work item type</param>
        /// <param name="direction">Field of the custom field of the single work item.</param>
        /// <param name="linkType">Type of the link of the single work item.</param>
        /// <param name="autoLinkType">Work item type of the work item to which to automatically link while publishing.</param>
        /// <param name="autoLinkSubtypeField">Field reference name of the field containing subtype of the automatically linked work item.</param>
        /// <param name="autoLinkSubtypeValue">Required subtype of of the work item to which to automatically link while publishing.</param>
        /// <param name="overwrite">Sets or gets whether the link overwrites existing links while publishing.</param>
        /// <returns>Configuration containing exactly one work item configuration.</returns>
        public static IConfiguration GetSimpleLinkConfiguration(string workItemType, Direction direction, string linkType, string autoLinkType, string autoLinkSubtypeField, string autoLinkSubtypeValue, bool overwrite)
        {
            var configuration = new Configuration();

            var item = new MappingElement
            {
                RelatedTemplate = "Generic10x1Table.xml",
                ImageFile = "standard.png",
                AssignRegularExpression = workItemType,
                Converters = new MappingConverter[0],
                WorkItemType = workItemType,
                MappingWorkItemType = workItemType,
                AssignCellCol = 1,
                AssignCellRow = 1,
                Fields = new[]
                {
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.ID", TableRow = 2, TableCol = 1 },
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.Rev", TableRow = 3, TableCol = 1 }
                },

                Links = new[]
                            {
                              new MappingLink { LinkValueType = linkType, AutomaticLinkWorkItemSubtypeField = autoLinkSubtypeField, AutomaticLinkWorkItemSubtypeValue = autoLinkSubtypeValue, AutomaticLinkWorkItemType = autoLinkType, Direction = direction, TableCol = 1, TableRow = 4, Overwrite = overwrite }
                            }
            };

            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, item, Path.Combine(Environment.CurrentDirectory, "Configuration")));
            return configuration;
        }

        /// <summary>
        /// Create a configuration with a subtype field
        /// </summary>
        /// <param name="templateName"></param>
        /// <param name="workItemType"></param>
        /// <param name="direction"></param>
        /// <param name="fieldValueType"></param>
        /// <param name="fieldName"></param>
        /// <param name="workItemSubtypeField"></param>
        /// <param name="workItemSubtypeValue"></param>
        /// <returns></returns>
        public static IConfiguration GetWorkItemSubtypeConfiguration(string templateName, string workItemType, Direction direction, FieldValueType fieldValueType, string fieldName, string workItemSubtypeField, string workItemSubtypeValue)
        {
            var configuration = new Configuration();

            var item = new MappingElement
            {
                RelatedTemplate = templateName,
                ImageFile = "standard.png",
                AssignRegularExpression = workItemType,
                Converters = new MappingConverter[0],
                WorkItemType = workItemType,
                MappingWorkItemType = workItemType,
                WorkItemSubtypeField = workItemSubtypeField,
                WorkItemSubtypeValue = workItemSubtypeValue,
                AssignCellCol = 1,
                AssignCellRow = 1,
                Fields = new[]
                {
                    new MappingField { Direction = Direction.TfsToOther, Name = "System.Id", TableRow = 2, TableCol = 3},
                    new MappingField { Direction = Direction.TfsToOther, Name ="System.Rev", TableRow = 2, TableCol = 4},
                    new MappingField { Direction = Direction.OtherToTfs, Name="System.Title", TableRow = 2, TableCol = 1}
                 },

                Links = new MappingLink[0]
            };

            configuration.GetConfigurationItems().Add(new ConfigurationItem(configuration, item, Path.Combine(Environment.CurrentDirectory, "Configuration")));
            return configuration;
        }



        /// <summary>
        /// Merges two configurations
        /// </summary>
        /// <param name="mergeInto">Configuration into which to merge.</param>
        /// <param name="mergeFrom">Configuration from which to merge.</param>
        /// <returns>Merged configuration containing both work item configurations.</returns>
        public static IConfiguration Merge(this IConfiguration mergeInto, IConfiguration mergeFrom)
        {
            if (mergeInto == null) throw new ArgumentNullException("mergeInto");
            if (mergeFrom == null) throw new ArgumentNullException("mergeFrom");

            foreach (var item in mergeFrom.GetConfigurationItems())
            {
                mergeInto.GetConfigurationItems().Add(item);
            }

            return mergeInto;
        }

    }
}
