#region Usings
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using AIT.TFS.SyncService.Model.Console;
using AIT.TFS.SyncService.Model.Helper;
using AIT.TFS.SyncService.Service.Configuration.Serialization.Console;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TFS.Test.Common;
// ReSharper disable InconsistentNaming
using System.Diagnostics;
#endregion

namespace TFS.SyncService.Test.Integration
{
    /// <summary>
    /// Integration test to test the call of the GetWorkItems command
    /// </summary>
    [TestClass]
    [DeploymentItem("Microsoft.WITDataStore32.dll")]
    [DeploymentItem("Microsoft.WITDataStore64.dll")]
    public class ConsoleExtensionHelperTest
    {
        #region Private Fields

        private static string Server = string.Empty; // see classinitialize
        private static string Project = string.Empty; // see classinitialize
        private const string TemplateName = "MSF For CMMI (2013)";
        private const bool Overwrite = false;
        private const bool CloseOnFinish = false;
        private static string WorkItemToImport = string.Empty; // see classinitialize
        private const string WrongWorkItemQuery = "CMMITemplate/SharedQueries/TestQuery";
        private const string RightWorkItemQuery = "WordToTFSTest_CMMI_TestReporting/Shared Queries/TestQuery";
        private static Application wordApplication;
        private static TestContext _testContext;

        #endregion Private Fields

        #region Test Initializations and Cleanup

        /// <summary>
        /// Create new document
        /// </summary>
        [TestInitialize]
        [TestCategory("Interactive")]
        public void TestInitialize()
        {
            TestCleanup.CloseWordDocumentAndKillOpenWordInstances();

        }

        ///// <summary>
        ///// Close all open word Documents after finishing the test
        ///// </summary>
        [TestCleanup]
        [TestCategory("Interactive")]
        public void CleanUp()
        {
            TestCleanup.CloseWordDocumentAndKillOpenWordInstances();
        }

        [ClassInitialize]
        public static void MyClassInitialize(TestContext testContext)
        {
            _testContext = testContext;
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);
            Server = CommonConfiguration.TfsTestServerConfiguration(_testContext).TeamProjectCollectionUrl;
            Project = CommonConfiguration.TfsTestServerConfiguration(_testContext).TeamProjectName;
            WorkItemToImport = CommonConfiguration.TfsTestServerConfiguration(_testContext).HtmlRequirementId.ToString();
        }

        /// <summary>
        /// Close all open word handles
        /// </summary>
        [ClassCleanup]
        [TestCategory("Interactive")]
        public static void ClassCleanup()
        {
            TestCleanup.CloseWordDocumentAndKillOpenWordInstances();
        }

        #endregion Test Initializations and Cleanup

        #region TestMethods

        /// <summary>
        /// Create document with arguments and config file, this should fail
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration/ConfigurationNew1.xml", "Config")]
        [ExpectedException(typeof(ArgumentException), "Einstellungen dürfen nur entweder über Kommandozeileneingabe oder über Konfigurationsdatei übergeben werden")]
        public void CreateBasicDocumentWithValidConfigFileAndValidSettings_ShouldReturnException()
        {
            //Create Doc conf with arguments and config file, this should fail
            // ReSharper disable once UnusedVariable
            var documentConfiguration = new DocumentConfiguration(Server, Project, GetTempDocumentNameAndCreateTempFolder(), TemplateName, Overwrite, CloseOnFinish, "Config/ConfigurationNew1.xml",false,null, TraceLevel.Verbose);
            Assert.Fail();
        }

        /// <summary>
        /// Create document with arguments and config file, this should fail
        /// </summary>
        [TestMethod]
        [DeploymentItem("Configuration/ConfigurationWithoutTestConfiguration.xml", "Config")]
         public void CreateDocumentConfigurationWithConfigFileWithoutTestSettings_ShouldCreateDocumentConfigurationObjectWithouException()
        {
            //Create Doc conf with arguments and config file, this should fail
            // ReSharper disable once UnusedVariable
            var documentConfiguration = new DocumentConfiguration(null, null, null, null, false, false, "Config/ConfigurationWithoutTestConfiguration.xml", false, null, TraceLevel.Verbose);

            Assert.IsNotNull(documentConfiguration);
        }

        /// <summary>
        /// Create document with arguments and config file, this should fail
        /// </summary>
        /// 
        [TestMethod]
        [DeploymentItem("Configuration/ConfigurationWithoutTestResultConfiguration.xml", "Config")]
        public void CreateDocumentConfigurationWithConfigFileWithoutTestResultConfiguration_ShouldCreateDocumentConfigurationObjectWithouException()
        {
            //Create Doc conf with arguments and config file, this should fail
            // ReSharper disable once UnusedVariable
            var documentConfiguration = new DocumentConfiguration(null, null, null, null, false, false, "Config/ConfigurationWithoutTestResultConfiguration.xml", false, null, TraceLevel.Verbose);

            Assert.IsNotNull(documentConfiguration);
        }

        /// <summary>
        /// Create document with arguments and config file, this should fail
        /// </summary>
        /// 
        [TestMethod]
        [DeploymentItem("Configuration/ConfigurationWithoutTestSpecificationConfiguration.xml", "Config")]
        public void CreateDocumentConfigurationWithConfigFileWithoutTestSpecificationConfiguration_ShouldCreateDocumentConfigurationObjectWithouException()
        {
            //Create Doc conf with arguments and config file, this should fail
            // ReSharper disable once UnusedVariable
            var documentConfiguration = new DocumentConfiguration(null, null, null, null, false, false, "Config/ConfigurationWithoutTestSpecificationConfiguration.xml", false, null, TraceLevel.Verbose);

            Assert.IsNotNull(documentConfiguration);
        }

        /// <summary>
        /// Test to create a basic document by id
        /// </summary>
        /// 
        [TestMethod]
        [DeploymentItem("Configuration/ConfigurationNew1.xml", "Config")]
        [TestCategory("Interactive2"), TestCategory("ConnectionNeeded")]
        public void CreateBasicDocumentByIdWithValidConfigFile_ShouldCreateDocument()
        {
            Assert.IsTrue(File.Exists("Config/ConfigurationNew1.xml"));
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);
            var fileName = GetTempDocumentNameAndCreateTempFolder();

            //Create Doc conf
            var documentConfiguration = new DocumentConfiguration(null, null, null, null, false, false, "Config/ConfigurationNew1.xml",false,null, TraceLevel.Verbose);

            documentConfiguration.Settings.Filename = fileName;

            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
                crm.CreateWorkItemDocument(documentConfiguration, WorkItemToImport, "ByID");

             //Wait for the import to finish
                while (crm.IsImporting)
                {
                    Thread.CurrentThread.Join(50);
                }
            Assert.IsTrue(File.Exists(fileName));
        }

        /// <summary>
        /// Test to create a basic document by id
        /// </summary>
        /// 
        [TestMethod]
        [DeploymentItem("Configuration/ConfigurationWithoutTestConfiguration.xml", "Config")]
        [TestCategory("Interactive2"), TestCategory("ConnectionNeeded")]
        public void CreateBasicDocumentByIdWithValidConfigFileWithoutTestConfiguratioNsettings_ShouldCreateDocument()
        {
            Assert.IsTrue(File.Exists("Config/ConfigurationWithoutTestConfiguration.xml"));
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);
            var fileName = GetTempDocumentNameAndCreateTempFolder();

            //Create Doc conf
            var documentConfiguration = new DocumentConfiguration(null, null, null, null, false, false, "Config/ConfigurationWithoutTestConfiguration.xml", false, null, TraceLevel.Verbose);

            documentConfiguration.Settings.Filename = fileName;

            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
            crm.CreateWorkItemDocument(documentConfiguration, WorkItemToImport, "ByID");

            //Wait for the import to finish
            while (crm.IsImporting)
            {
                Thread.CurrentThread.Join(50);
            }
            Assert.IsTrue(File.Exists(fileName));
        }

        /// <summary>
        /// Test that makes sure that the connection information is saved
        /// </summary>
        /// 
        [TestMethod]
        [DeploymentItem("Configuration\\ConfigurationNew1.xml", "Config")]
        [TestCategory("Interactive2"), TestCategory("ConnectionNeeded")]
        public void CreateBasicDocumentByIdWithValidConfigFile_EnsureThatConnectionInformationisSavedInDocumentMode_ShouldCreateDocumentWithValidConnectionInfo()
        {
            Assert.IsTrue(File.Exists("Config\\ConfigurationNew1.xml"));
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);
            var fileName = GetTempDocumentNameAndCreateTempFolder();

            //Create Doc conf
            var documentConfiguration = new DocumentConfiguration(null, null, null, null, false, false, "Config\\ConfigurationNew1.xml", false, null, TraceLevel.Verbose);

            documentConfiguration.Settings.Filename = fileName;

            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
            crm.CreateWorkItemDocument(documentConfiguration, WorkItemToImport, "ByID");

            //Wait for the import to finish
            Assert.IsTrue(File.Exists(fileName));
        }

        /// <summary>
        /// Test to create a basic document by id
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive2"), TestCategory("ConnectionNeeded")]
        public void CreateBasicDocumentByIdWithStandardOptions_ShouldCreateDocument()
        {
            var fileName = GetTempDocumentNameAndCreateTempFolder();
            //Create Doc conf
            var documentConfiguration = new DocumentConfiguration(Server, Project, fileName, TemplateName, Overwrite, CloseOnFinish, null,false,null, TraceLevel.Verbose);
            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
            crm.CreateWorkItemDocument(documentConfiguration, WorkItemToImport, "ByID");

            //Wait for the import to finish
            while (crm.IsImporting)
            {
                Thread.CurrentThread.Join(50);
            }

            Assert.IsTrue(File.Exists(fileName));
            Assert.IsTrue(SearchActiveWordDocumentForString(WorkItemToImport));
        }

        /// <summary>
        /// Test to create a basic document by id
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive2"), TestCategory("ConnectionNeeded")]
        [ExpectedException(typeof(Exception))]
        public void CreateBasicDocumentByQuerydWithStandardOptionsAndWrongQueryName_ShouldFail()
        {
            var fileName = GetTempDocumentNameAndCreateTempFolder();
            //Create Doc conf
            var documentConfiguration = new DocumentConfiguration(Server, Project, fileName, TemplateName, Overwrite, CloseOnFinish, null,false,null, TraceLevel.Verbose);
            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
            crm.CreateWorkItemDocument(documentConfiguration, WrongWorkItemQuery, "ByQuery");
            Assert.IsFalse(File.Exists(fileName));
        }

        /// <summary>
        /// Test to create a basic document by id
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive2")]
        [TestCategory("ConnectionNeeded")]
        public void CreateBasicDocumentByQuerydWithStandardOptionsAndCorrectQueryName_ShouldCreateDocument()
        {
            var fileName = GetTempDocumentNameAndCreateTempFolder();
            //Create Doc conf
            var documentConfiguration = new DocumentConfiguration(Server, Project, fileName, TemplateName, Overwrite, CloseOnFinish, null,true,null, TraceLevel.Verbose);
            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
            crm.CreateWorkItemDocument(documentConfiguration, RightWorkItemQuery, "ByQuery");

            //Wait for the import to finish
            while (crm.IsImporting)
            {
                Thread.CurrentThread.Join(50);
            }

            Assert.IsTrue(File.Exists(fileName));
            Assert.IsTrue(SearchActiveWordDocumentForString("22"));
            Assert.IsTrue(SearchActiveWordDocumentForString("26"));

            Assert.IsTrue(SearchActiveWordDocumentForString("Affected Requirement"));
            Assert.IsTrue(SearchActiveWordDocumentForString("Parent Requirement"));
        }

        /// <summary>
        /// Test to create a basic document for an existing document
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        [ExpectedException(typeof(Exception), "File Exists")]
        public void CreateDocumentContainingWorkItemsForExistingDocument_OverwriteSwitchedOff_ShouldFail()
        {
            var filename = GetTempDocumentNameAndCreateTempFolder();
            var newWordApp = new Application();
            //Create a document, save it and close it
            newWordApp.Documents.Add();
            var currentDocument = newWordApp.ActiveDocument;
            currentDocument.SaveAs(filename);
            currentDocument.Close();
            newWordApp.Quit();

            Assert.IsTrue(File.Exists(filename));

            var documentConfiguration = new DocumentConfiguration(Server, Project, filename, TemplateName, false, CloseOnFinish, null,false, null, TraceLevel.Verbose);
            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
            crm.CreateWorkItemDocument(documentConfiguration, WorkItemToImport, "ByID");

            //Wait for the import to finish
            while (crm.IsImporting)
            {
                Thread.CurrentThread.Join(50);
            }

            Assert.IsTrue(SearchActiveWordDocumentForString(WorkItemToImport));

            Assert.IsFalse(SearchSpecificWordDocumentForString(WorkItemToImport, filename));
            //File should exist anyway
            Assert.IsTrue(File.Exists(filename));
        }

        /// <summary>
        /// Test to create a basic document
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive2"), TestCategory("ConnectionNeeded")]
        public void CreateDocumentContainingWorkItemsForExistingDocument_OverwriteSwitchedOn_ContentOfDocumentShouldBeReplaced()
        {
            var filename = GetTempDocumentNameAndCreateTempFolder();
            var newWordApp = new Application();
            //Create a document, save it and close it
            newWordApp.Documents.Add();
            var currentDocument = newWordApp.ActiveDocument;
            currentDocument.SaveAs(filename);
            currentDocument.Close();
            newWordApp.Quit();

            Assert.IsTrue(File.Exists(filename));

            var documentConfiguration = new DocumentConfiguration(Server, Project, filename, TemplateName, true, CloseOnFinish, null,false, null, TraceLevel.Verbose);
            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
            crm.CreateWorkItemDocument(documentConfiguration, WorkItemToImport, "ByID");

            //Wait for the Importing to finish
            while (crm.IsImporting)
            {
                Thread.CurrentThread.Join(50);
            }

            Assert.IsTrue(SearchActiveWordDocumentForString(WorkItemToImport));

            //The id should also be in the saved file
            Assert.IsTrue(SearchSpecificWordDocumentForString(WorkItemToImport, filename));
            //File should exist anyway
            Assert.IsTrue(File.Exists(filename));
        }

        /// <summary>
        /// Test to create a document based on a dotx template file
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive2"), TestCategory("ConnectionNeeded")]        
        [DeploymentItem("Configuration/WordTemplate.dotx", "Configuration")]
        [DeploymentItem("Configuration/ConfigurationNew1.xml", "Configuration")]
        public void CreateBasicDocumentByIdWithValidSettingsAndTemplate_ShouldCreateDocumentThatContainsWorkItemAndTextFromWordTempalte()
        {
            Assert.IsTrue(File.Exists("Configuration/WordTemplate.dotx"));
            Assert.IsTrue(File.Exists("Configuration/ConfigurationNew1.xml"));
            var fileName = GetTempDocumentNameAndCreateTempFolder();
            //Create Doc conf
            var documentConfiguration = new DocumentConfiguration(Server, Project, fileName, TemplateName, true, CloseOnFinish, null, false, "Configuration/WordTemplate.dotx", TraceLevel.Verbose);

            documentConfiguration.Settings.Filename = fileName;
            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
            crm.CreateWorkItemDocument(documentConfiguration, WorkItemToImport, "ByID");

            //Wait for the Importing to finish
            while (crm.IsImporting)
            {
                Thread.CurrentThread.Join(50);
            }

              Assert.IsTrue(File.Exists(fileName));
              Assert.IsTrue(SearchSpecificWordDocumentForString(WorkItemToImport, fileName));
              Assert.IsTrue(SearchSpecificWordDocumentForString("This is a word template file.", fileName));
        }

        /// <summary>
        /// Testing the exeception that occurs if a tempalte file does not exist
        /// </summary>
        /// 
        [TestMethod]
        [DeploymentItem("Configuration\\WordTemplate.dotx", "Configuration")]
        [DeploymentItem("Configuration\\ConfigurationNew1.xml", "Configuration")]
        [ExpectedException(typeof(Exception))]
        [TestCategory("Interactive")]
        public void CreateBasicDocumentByIdWithValidSettingsAndMissingTemplate_ShouldNotCreateDocumentAndShoudlRaiseException()
        {
            Assert.IsTrue(File.Exists("Configuration\\WordTemplate.dotx"));
            Assert.IsTrue(File.Exists("Configuration\\ConfigurationNew1.xml"));
            var fileName = GetTempDocumentNameAndCreateTempFolder();

            //Create Doc conf with wrong reference
            var documentConfiguration = new DocumentConfiguration(Server, Project, fileName, TemplateName, true, CloseOnFinish, null, false, "Configuration/NONExistingWordTemplate.dotx", TraceLevel.Verbose);

            documentConfiguration.Settings.Filename = fileName;
            var crm = new ConsoleExtensionHelper(new TestReportingProgressCancellationService(false));
            crm.CreateWorkItemDocument(documentConfiguration, WorkItemToImport, "ByID");

            Assert.IsFalse(File.Exists(fileName));

        }

        #endregion TestMethods

        #region TestHelpers

        /// <summary>
        /// Search a given word document for a string
        /// </summary>
        /// <param name="searchString"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private bool SearchSpecificWordDocumentForString(string searchString, string fileName)
        {
            //Create a document, save it and close it
            GetCurrentWordApp();

            object miss = System.Reflection.Missing.Value;
            object path = fileName;
            object readOnly = true;
            Document docs = wordApplication.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            string totaltext = "";
            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                totaltext += " \r\n " + docs.Paragraphs[i + 1].Range.Text;
            }
            Console.WriteLine(totaltext);
            docs.Close();

            return totaltext.Contains(searchString);
        }

        /// <summary>
        /// Search the active word document for a string
        /// </summary>
        /// <param name="searchString"></param>
        /// <returns></returns>
        private bool SearchActiveWordDocumentForString(string searchString)
        {
            GetCurrentWordApp();
            string totaltext = "";
            for (int i = 0; i < wordApplication.ActiveDocument.Paragraphs.Count; i++)
            {
                totaltext += " \r\n " + wordApplication.ActiveDocument.Paragraphs[i + 1].Range.Text;
            }

            return totaltext.Contains(searchString);
        }

        /// <summary>
        /// Helper to create a temp folder and to create a temp file in the temp path
        /// </summary>
        /// <returns></returns>
        private string GetTempDocumentNameAndCreateTempFolder()
        {
            string tempPath = Path.GetTempPath();
            Guid folderGuid = Guid.NewGuid();
            Guid fileGuid = Guid.NewGuid();
            Directory.CreateDirectory(Path.Combine(tempPath, folderGuid.ToString()));


            var di = new DirectoryInfo(Path.Combine(tempPath, folderGuid.ToString()));
            di.Attributes &= ~FileAttributes.ReadOnly;

            string filename = Path.Combine(tempPath, folderGuid.ToString(), fileGuid.ToString() + ".docx");
            return filename;
        }

        /// <summary>
        /// Get the current word app
        /// </summary>
        private void GetCurrentWordApp()
        {
            try
            {
                wordApplication = (Application)Marshal.GetActiveObject("Word.Application");
            }
            catch (COMException)
            {
                Type type = Type.GetTypeFromProgID("Word.Application");
                wordApplication = (Application)Activator.CreateInstance(type);
            }
        }

        #endregion TestHelpers
    }
}
