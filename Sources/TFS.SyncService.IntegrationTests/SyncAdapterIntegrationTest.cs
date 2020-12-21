#region Usings
using System;
using System.Collections.Generic;
using System.Linq;
using AIT.TFS.SyncService.Adapter.TFS2012;
using AIT.TFS.SyncService.Adapter.Word2007;
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Service;
using Microsoft.Office.Interop.Word;
using Microsoft.TeamFoundation.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TFS.Test.Common;
using AIT.TFS.SyncService.Contracts.Configuration;
using Moq;
// ReSharper disable InconsistentNaming
#endregion

namespace TFS.SyncService.Test.Integration
{
    /// <summary>
    /// Integration test for WordToTFS
    /// The combination of how word exports HTML, TFS saves HTML and how HTML are compared cannot reasonably be tested in isolation.
    /// </summary>
    [TestClass]
    public class SyncAdapterIntegrationTest
    {
        private Document _document;
        private bool _ignoreFailOnUserInforation;
        private static TestContext _testContext;

        #region Test Initializations and Cleanup

        /// <summary>
        /// Make sure all services are registered.
        /// </summary>
        [ClassInitialize]
        [TestCategory("Interactive")]
        public static void MyClassInitialize(TestContext testContext)
        {
            _testContext = testContext;
            CommonConfiguration.ReplaceConfigFileTokens(_testContext);

            // ReSharper disable once UnusedVariable
            var projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(CommonConfiguration.TfsTestServerConfiguration(_testContext).TeamProjectCollectionUrl));

            AIT.TFS.SyncService.Service.AssemblyInit.Instance.Init();
        }

        [TestInitialize]
        [TestCategory("Interactive")]
        public void TestInitialize()
        {
            _ignoreFailOnUserInforation = false;
            _document = new Document();
        }

        [TestCleanup]
        [TestCategory("Interactive")]
        public void TestCleanup()
        {
            if (!_ignoreFailOnUserInforation)
                FailOnUserInformation();

            _document?.Close(false);
        }

        #endregion Test Initializations and Cleanup

        #region TestMethods

        [TestMethod]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void IntegrationTest_HTMLField_GetPublish_ShouldNotIncreaseRevision()
        {
            var serverConfig = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            serverConfig.Configuration.IgnoreFormatting = true;

            // import
            var importItems = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            var oldRevision = importItems.Single().Revision;
            Refresh(serverConfig, importItems);

            // publish
            var publishWorkItems = GetWordWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            Publish(serverConfig, publishWorkItems, false);

            // get published item
            var publishedItems = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            var newRevision = publishedItems.Single().Revision;

            Assert.IsTrue(publishedItems.Single().Fields[serverConfig.HtmlFieldReferenceName].Value.Contains("<"), "Ait.Common.Description must contain html data.");
            Assert.AreEqual(oldRevision, newRevision, "Revision changed after Get/Publish");
        }

        [TestMethod]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void IntegrationTest_HTMLField_PublishTwice_ShouldNotIncreaseRevisionSecondTime()
        {
            var serverConfig = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            serverConfig.Configuration.IgnoreFormatting = true;

            // import
            var importItems = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            Refresh(serverConfig, importItems);

            // change
            _document.Tables[1].Cell(3, 1).Range.Text = Guid.NewGuid().ToString();
            _document.Tables[1].Cell(3, 1).Range.Font.Color = WdColor.wdColorDarkRed;

            // publish
            var publishWorkItems = GetWordWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            Publish(serverConfig, publishWorkItems, false);
            var firstRevision = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).Single().Revision;
            Publish(serverConfig, publishWorkItems, false);
            var secondRevision = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).Single().Revision;

            Assert.AreEqual(firstRevision, secondRevision, "Revision changed after second publish.");
        }

        [TestMethod]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void IntegrationTest_HTMLField_Refresh_MustNotOverwriteDifferentWordFormatting()
        {
            var serverConfig = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            serverConfig.Configuration.IgnoreFormatting = true;

            // import
            var importItems = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            Refresh(serverConfig, importItems);

            // change
            _document.Tables[1].Cell(3, 1).Range.Font.Color = WdColor.wdColorDarkGreen;

            // refresh
            Refresh(serverConfig, GetWordWorkItems(null, serverConfig));

            Assert.AreEqual(WdColor.wdColorDarkGreen, _document.Tables[1].Cell(3, 1).Range.Font.Color);
        }

        [TestMethod]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void IntegrationTest_HTMLField_Refresh_ShouldRefreshWordFormattingIfContentIsDifferent()
        {
            var serverConfig = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            serverConfig.Configuration.IgnoreFormatting = true;

            // import
            var importItems = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            Refresh(serverConfig, importItems);

            // change
            _document.Tables[1].Cell(3, 1).Range.Text = "different";
            _document.Tables[1].Cell(3, 1).Range.Font.Color = WdColor.wdColorDarkGreen;

            // refresh
            Refresh(serverConfig, GetWordWorkItems(null, serverConfig));

            Assert.AreNotEqual(WdColor.wdColorDarkGreen, _document.Tables[1].Cell(3, 1).Range.Font.Color);
        }

        /// <summary>
        /// Check if the revisionsettings stay the same
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        public void IntegrationTest_TrackChanges_Publish_ShouldDeactivteTrackChanges()
        {
            var serverConfig = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            serverConfig.Configuration.IgnoreFormatting = true;

            // import
            var importItems = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            Refresh(serverConfig, importItems);

            //Set Track changes to true
            bool trackRevisionsBefore = _document.TrackRevisions;

            // change
            _document.Tables[1].Cell(3, 1).Range.Text = Guid.NewGuid().ToString();
            _document.Tables[1].Cell(3, 1).Range.Font.Color = WdColor.wdColorDarkRed;

            // publish
            var publishWorkItems = GetWordWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            Publish(serverConfig, publishWorkItems, false);

            //Revisions should be the same after publish
            Assert.AreEqual(_document.TrackRevisions, trackRevisionsBefore);
        }

        [TestMethod]
        [TestCategory("Interactive")]
        [TestCategory("ConnectionNeeded")]
        [Ignore] // Cannot test image behaviour because word does not have access to the attachment handler of TFS2012/TFS2010 an therefore cannot insert images from tfs work items.
        public void IntegrationTest_HTMLField_RefreshChangedImageNoText_ShouldReplaceWordImage()
        {
            /*
            var serverConfig = CommonConfiguration.Tfs2012ServerConfiguration;
            serverConfig.Configuration.IgnoreFormatting = true;

            // import work item
            var importItems = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();
            Refresh(serverConfig, importItems);

            // make sure tfs work item has a different image
            var tfsAdapter = new Tfs2012SyncAdapter(serverConfig.TeamProjectCollectionUrl, serverConfig.TeamProjectName, null, serverConfig.Configuration);
            tfsAdapter.Open(new[] { serverConfig.HtmlRequirementId });
            tfsAdapter.WorkItems.Single().Fields[serverConfig.HTMLFieldReferenceName].Value = string.Format("<img src=\"{0}\"/>", Path.GetFullPath("Configuration\\Image1.jpg"));
            Assert.IsFalse(tfsAdapter.Save().Any());

            // replace description
            _document.Tables[1].Cell(3, 1).Range.InlineShapes.AddPicture(Path.GetFullPath("Configuration\\Image2.jpg"));
            var height = _document.Tables[1].Cell(3, 1).Range.InlineShapes[1].Height;
            _document.Tables[1].Cell(3, 1).Range.Collapse();

            // refresh 
            Refresh(serverConfig, GetWordWorkItems(null, serverConfig));

            Assert.AreNotEqual(height, _document.Tables[1].Cell(3, 1).Range.InlineShapes[1].Height, 1);
             */

          Assert.Inconclusive("Cannot test image behaviour because word does not have access to the attachment handler of TFS2012/TFS2010 an therefore cannot insert images from tfs work items.");
        }

        ///  <summary>
        /// Test if a workitem that contains unclosed lists is published. This should not be the case as an error is thrown by WordToTFS
        ///</summary>
        [TestMethod]        
        [TestCategory("Interactive2"), TestCategory("ConnectionNeeded")]
        [DeploymentItem("Configuration", "Configuration")]
        [DeploymentItem("Microsoft.WITDataStore32.dll")]
        [DeploymentItem("Microsoft.WITDataStore64.dll")]
        public void IntegrationTest_PublishOfWorkItemWithUnclosedList_ShouldNotBePublished()
        {
            var systemVariableMock = new Mock<ISystemVariableService>();
            _ignoreFailOnUserInforation = true;

            var serverConfig = CommonConfiguration.TfsTestServerConfiguration(_testContext);
            serverConfig.Configuration.IgnoreFormatting = false;
            var requirementConf = serverConfig.Configuration.GetWorkItemConfiguration("Requirement");

            // import an workitems to use the structure as example for the new workitem
            var importItems = GetTfsWorkItems(new[] { serverConfig.HtmlRequirementId }, serverConfig).ToList();

            Refresh(serverConfig, importItems);

            //Add Bullet points to the table.
             _document.Tables[1].Cell(3, 1).Range.ListFormat.ApplyBulletDefault();
            var bulletItems = new[] { "One", "Two", "Three" };

            for (int i = 0; i < bulletItems.Length; i++)
            {
                string bulletItem = bulletItems[i];
                if (i < bulletItems.Length - 1)
                    bulletItem = bulletItem + "\n";
               _document.Tables[1].Cell(3, 1).Range.InsertBefore(bulletItem);
            }

            //delete the id
            _document.Tables[1].Cell(2, 3).Range.Text = string.Empty;
            _document.Tables[1].Cell(2, 1).Range.Text = "TestTitle";

            var createdWorkItem = new WordTableWorkItem(_document.Tables[1], "Requirement", serverConfig.Configuration, requirementConf);

            // publish
           var publishWorkItems = new List<IWorkItem>();
           publishWorkItems.Add(createdWorkItem);

           var configuration = serverConfig.Configuration;
           var tfsAdapter = new Tfs2012SyncAdapter(serverConfig.TeamProjectCollectionUrl, serverConfig.TeamProjectName, null, configuration);
           var wordAdapter = new Word2007TableSyncAdapter(_document, configuration);

           var syncService = new WorkItemSyncService(systemVariableMock.Object);
           syncService.Publish(wordAdapter, tfsAdapter, publishWorkItems, false, configuration);

            //Get the information storage
           var informationStorage = SyncServiceFactory.GetService<IInfoStorageService>();
           var error = informationStorage.UserInformation.FirstOrDefault(x => x.Type == UserInformationType.Error);

            //Test that an error was thrown
            Assert.IsNotNull(error);

            //Clear the information storage
            informationStorage.UserInformation.Clear();
           }

        #endregion

        #region TestHelpers
        /// <summary>
        /// Publish the information to tfs
        /// </summary>
        /// <param name="serverConfiguration"></param>
        /// <param name="workItems"></param>
        /// <param name="forceOverwrite"></param>
        private void Publish(ServerConfiguration serverConfiguration, IEnumerable<IWorkItem> workItems, bool forceOverwrite)
        {
            var systemVariableMock = new Mock<ISystemVariableService>();
            var configuration = serverConfiguration.Configuration;
            var tfsAdapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, configuration);
            var wordAdapter = new Word2007TableSyncAdapter(_document, configuration);

            var syncService = new WorkItemSyncService(systemVariableMock.Object);
            syncService.Publish(wordAdapter, tfsAdapter, workItems, forceOverwrite, configuration);
            FailOnUserInformation();
        }

        /// <summary>
        /// Helper method to refresh information
        /// </summary>
        /// <param name="serverConfiguration"></param>
        /// <param name="workItems"></param>
        private void Refresh(ServerConfiguration serverConfiguration, IEnumerable<IWorkItem> workItems)
        {
            var systemVariableMock = new Mock<ISystemVariableService>();
            var configuration = serverConfiguration.Configuration;
            var tfsAdapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, configuration);
            var wordAdapter = new Word2007TableSyncAdapter(_document, configuration);

            var syncService = new WorkItemSyncService(systemVariableMock.Object);
            syncService.Refresh(tfsAdapter, wordAdapter, workItems, configuration);
            FailOnUserInformation();
        }

        /// <summary>
        /// Helper method to get a work item from word
        /// </summary>
        /// <param name="ids"></param>
        /// <param name="serverConfiguration"></param>
        /// <returns></returns>
        private IEnumerable<IWorkItem> GetWordWorkItems(int[] ids, ServerConfiguration serverConfiguration)
        {
            var configuration = serverConfiguration.Configuration;
            var wordAdapter = new Word2007TableSyncAdapter(_document, configuration);
            wordAdapter.Open(ids);
            return wordAdapter.WorkItems;
        }

        /// <summary>
        /// Helper method to get work items from tfs
        /// </summary>
        /// <param name="ids"></param>
        /// <param name="serverConfiguration"></param>
        /// <returns></returns>
        private IEnumerable<IWorkItem> GetTfsWorkItems(int[] ids, ServerConfiguration serverConfiguration)
        {
            var configuration = serverConfiguration.Configuration;
            var tfsAdapter = new Tfs2012SyncAdapter(serverConfiguration.TeamProjectCollectionUrl, serverConfiguration.TeamProjectName, null, configuration);

            tfsAdapter.Open(ids);
            return tfsAdapter.WorkItems;
        }

        /// <summary>
        /// Helper method that catches exceptions and user information raised during publish / refresh
        /// </summary>
        private void FailOnUserInformation()
        {
            var informationStorage = SyncServiceFactory.GetService<IInfoStorageService>();
            var error = informationStorage.UserInformation.FirstOrDefault(x => x.Type == UserInformationType.Error);
            if (error != null)
            {
                Assert.Fail(error.Explanation);
            }
        }
        #endregion
    }
}
