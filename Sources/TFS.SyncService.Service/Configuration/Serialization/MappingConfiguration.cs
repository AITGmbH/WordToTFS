using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
// TODO What is the purpose of the DefaultValue attribute here? -SER

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    /// <summary>
    /// Class used as top serialization class to serialize all mapping settings.
    /// </summary>
    [XmlRoot("MappingConfiguration", Namespace = "", IsNullable = false)]
    [XmlType("MappingConfiguration", Namespace = "")]
    public class MappingConfiguration
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="MappingConfiguration"/> class.
        /// </summary>
        public MappingConfiguration()
        {
            ShowMappedCellNotExistsMessage = true;
            ShowMappedFieldNotExistsMessage = true;
            EnablePublish = true;
            EnableRefresh = false;
            EnableGetWorkItems = true;
            EnableOverview = true;
            EnableHistoryComment = false;
            EnableDeleteIds = true;
            EnableEmpty = true;
            EnableNew = true;
            EnableEditDefaultValues = true;
            EnableAreaIterationPath = true;
            IgnoreFormatting = false;
            ConflictOverwrite = false;
            EnableConflictOverwriteSwitch = true;
            EnableIgnoreFormattingSwitch = true;
            HideElementInWord = false;
            TypeOfHierarchyRelationships = "";
            TypeOfHierachyRelationships = "";
            Variables = new List<Variable>();
            ObjectQueries = new List<ObjectQuery>();

            DefaultProjectName = "";
            DefaultServerUrl = "";
            AutoConnect = false;

            EnableTemplateManager = true;
            EnableTemplateSelection = true;
            GetDirectLinksOnly = false;

            AutoRefreshQuery = "";
        }


        #endregion Constructors

        #region Public serialization properties

        /// <summary>
        /// Property used to serialize 'ShowName' xml attribute.
        /// </summary>
        [XmlAttribute("ShowName")]
        public string ShowName { get; set; }

        /// <summary>
        /// Property used to serialize 'RelatedSchema' xml attribute.
        /// </summary>
        [XmlAttribute("RelatedSchema")]
        public string RelatedSchema { get; set; }

        /// <summary>
        /// Property used to serialize 'Variables' xml attirbute
        /// </summary>
        [XmlArray("Variables")]
        [XmlArrayItem("Variable")]
        public List<Variable> Variables{get;set;}


        /// <summary>
        /// 
        /// </summary>
        [XmlArray("ObjectQueries")]
        [XmlArrayItem("ObjectQuery")]
        public List<ObjectQuery> ObjectQueries
        {
            get;
            set;
        }

        /// <summary>
        /// Property used to serialize 'DefaultMapping' xml attribute.
        /// </summary>
        [XmlAttribute("DefaultMapping")]
        public bool DefaultMapping { get; set; }


        /// <summary>
        /// Property used to serialize 'ButtonsCustomization' xml nodes.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
        [XmlArray("ButtonsCustomization")]
        [XmlArrayItem("Button")]
        public List<ButtonCustomization> Buttons { get; set; }

        /// <summary>
        /// Property used to serialize "Mappings' xml node.
        /// </summary>
        [XmlElement("Mappings")]
        public MappingList MappingList { get; set; }

        /// <summary>
        /// Property used to serialize "Headers' xml node.
        /// </summary>
        [XmlArray("Headers")]
        [XmlArrayItem("Header")]
        public List<MappingHeader> Headers { get; set; }

        /// <summary>
        /// Gets a value indicating whether the message for the specific problem should be showed.
        /// </summary>
        [XmlAttribute("ShowMappedCellNotExistsMessage")]
        [DefaultValue(true)]
        public bool ShowMappedCellNotExistsMessage { get; set; }

        /// <summary>
        /// Gets a value indicating whether the message for the specific problem should be showed.
        /// </summary>
        [XmlAttribute("ShowMappedFieldNotExistsMessage")]
        [DefaultValue(true)]
        public bool ShowMappedFieldNotExistsMessage { get; set; }

        /// <summary>
        /// Property serialize "UseStackRank" xml node
        /// </summary>
        [XmlAttribute("UseStackRank")]
        [DefaultValue(false)]
        public bool UseStackRank { get; set; }

        /// <summary>
        /// Property serialize "EnableRefresh" xml node
        /// </summary>
        [XmlAttribute("EnableRefresh"), DefaultValue(false)]

        public bool EnableRefresh { get; set; }

        /// <summary>
        /// Property serialize "EnablePublish" xml node
        /// </summary>
        [XmlAttribute("EnablePublish"), DefaultValue(true)]
        public bool EnablePublish { get; set; }

        /// <summary>
        /// Gets whether the checkbox that enables format ignoring is visible
        /// </summary>
        [XmlAttribute("EnableIgnoreFormattingSwitch")]
        [DefaultValue(true)]
        public bool EnableIgnoreFormattingSwitch { get; set; }

        /// <summary>
        /// Gets whether the checkbox that enables overwriting of conflicting items is visible
        /// </summary>
        [XmlAttribute("EnableConflictOverwriteSwitch")]
        [DefaultValue(true)]
        public bool EnableConflictOverwriteSwitch { get; set; }

        /// <summary>
        /// Property serialize "EnableGetWorkItems" xml node
        /// </summary>
        [XmlAttribute("EnableGetWorkItems"), DefaultValue(true)]
        public bool EnableGetWorkItems { get; set; }

        /// <summary>
        /// Property serialize "EnableOverview" xml node
        /// </summary>
        [XmlAttribute("EnableOverview"), DefaultValue(true)]
        public bool EnableOverview { get; set; }

        /// <summary>
        /// Property serialize "EnableHistoryComment" xml node
        /// </summary>
        [XmlAttribute("EnableHistoryComment"), DefaultValue(false)]
        public bool EnableHistoryComment { get; set; }

        /// <summary>
        /// Property serialize "EnableEmpty" xml node
        /// </summary>
        [XmlAttribute("EnableEmpty"), DefaultValue(true)]
        public bool EnableEmpty { get; set; }

        /// <summary>
        /// Property serialize "EnableNew" xml node
        /// </summary>
        [XmlAttribute("EnableNew"), DefaultValue(true)]
        public bool EnableNew { get; set; }

        /// <summary>
        /// Property serialize "EnableDeleteIds" xml node
        /// </summary>
        [XmlAttribute("EnableDeleteIds"), DefaultValue(true)]
        public bool EnableDeleteIds { get; set; }

        /// <summary>
        /// Property serialize "EnableEditDefaultValues" xml node
        /// </summary>
        [XmlAttribute("EnableEditDefaultValues"), DefaultValue(true)]
        public bool EnableEditDefaultValues { get; set; }

        /// <summary>
        /// Property serialize "EnableAreaIterationPath" xml node
        /// </summary>
        [XmlAttribute("EnableAreaIterationPath"), DefaultValue(true)]
        public bool EnableAreaIterationPath { get; set; }

        /// <summary>
        /// Property serialize "IgnoreFormatting" xml node
        /// </summary>
        [XmlAttribute("IgnoreFormatting")]
        [DefaultValue(true)]
        public bool IgnoreFormatting { get; set; }

        /// <summary>
        /// Property serialize "IgnoreFormatting" xml node.
        /// </summary>
        [XmlAttribute("ConflictOverwrite")]
        [DefaultValue(false)]
        public bool ConflictOverwrite { get; set; }

        /// <summary>
        /// Property serialize "HideElementInWord" xml node.
        /// </summary>
        [XmlAttribute("HideElementInWord"),DefaultValue(false)]
        public bool HideElementInWord {get; set;}

        /// <summary>
        /// Property serialize "TypeOfHierarchyRelationships" xml node.
        /// This is the corrected Version from Version 4.1
        /// </summary>
        [XmlAttribute("TypeOfHierarchyRelationships"), DefaultValue("")]
        public string TypeOfHierarchyRelationships { get; set; }

        /// <summary>
        /// Property serialize "TypeOfHierachyRelationships" xml node.
        /// This is the misspelled property form version 4.0
        /// </summary>
        [XmlAttribute("TypeOfHierachyRelationships"), DefaultValue("")]
        public string TypeOfHierachyRelationships { get; set; }

        /// <summary>
        /// Property serialize "AttachmentFolderMode" xml node.
        /// </summary>
        [XmlAttribute("AttachmentFolderMode"), DefaultValue(AttachmentFolderMode.WithGuid)]
        public AttachmentFolderMode AttachmentFolderMode { get; set; }

        /// <summary>
        /// Property used to serialize pre-operations.
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
        [XmlArray("PreOperations")]
        [XmlArrayItem("Operation")]
        public List<OperationConfiguration> PreOperations { get; set; }

        /// <summary>
        /// Property used to serialize post-operations
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Xml serialization")]
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly", Justification = "Xml serialization")]
        [XmlArray("PostOperations")]
        [XmlArrayItem("Operation")]
        public List<OperationConfiguration> PostOperations { get; set; }

        /// <summary>
        /// Property used to show query tree collapsed or expanded.
        /// </summary>
        [XmlAttribute("CollapsQueryTree")]
        [DefaultValue(false)]
        public bool CollapsQueryTree { get; set; }

        #endregion Public serialization properties

        #region Public serialization properties - test related

        /// <summary>
        /// Gets or sets the test configuration part.
        /// </summary>
        [XmlElement("TestConfiguration")]
        public TestConfiguration TestConfiguration { get; set; }

        /// <summary>
        /// Property to specific a server url for the auto connect functionality
        /// </summary>
        [XmlAttribute("DefaultServerUrl")]
        public string DefaultServerUrl
        {
            get;
            set;
        }

        /// <summary>
        /// Property to specific a project name for the auto connect functionality
        /// </summary>
        [XmlAttribute("DefaultProjectName")]
        public string DefaultProjectName
        {
            get;
            set;
        }

        /// <summary>
        /// Property to specific a project name for the auto connect functionality
        /// </summary>
        [XmlAttribute("AutoConnect")]
        [DefaultValue(false)]
        public bool AutoConnect
        {
            get;
            set;
        }

        /// <summary>
        /// Prooperty to enable or to disable the template Manager
        /// </summary>
        [XmlAttribute("EnableTemplateManager"), DefaultValue(true)]
        public bool EnableTemplateManager
        {
            get;
            set;
        }

        /// <summary>
        /// Property to enable or to disable the template Manager
        /// </summary>
        [XmlAttribute("EnableTemplateSelection"),
        DefaultValue(true)]
        public bool EnableTemplateSelection
        {
            get;
            set;
        }

        /// <summary>
        /// Determines whether to get direct links only or recursive
        /// </summary>
        [XmlAttribute("GetDirectLinksOnly"),
        DefaultValue(true)]
        public bool GetDirectLinksOnly
        {
            get;
            set;
        }


        /// <summary>
        /// Name of a property that is refreshed automaticly on connect
        /// </summary>
        [XmlAttribute("AutoRefreshQuery")]
        public string AutoRefreshQuery
        {
            get;
            set;
        }

        #endregion Public serialization properties - test related
    }
}