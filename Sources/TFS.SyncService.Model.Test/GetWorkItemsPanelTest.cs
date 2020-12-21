#region Usings
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Word;
using AIT.TFS.SyncService.Model.WindowModel;
using System.Collections.Generic;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using System.Linq;
#endregion

namespace TFS.SyncService.Model.Test.Unit
{
    [TestClass]
    public class GetWorkItemsPanelTest
    {
        #region TestMethods
        [TestMethod]
        public void TestCompareCountsOfWorkItems_ShouldReturnDictionary()
        {
            // Arange
            var moqtfsService = new Mock<ITfsService>();
            moqtfsService.Setup(x => x.GetLinkTypes()).Returns((ICollection<ITFSWorkItemLinkType>)null);
            var moqConfiguration = new Mock<IConfiguration>();
            var moqSyncServiceDocumentModel = new Mock<ISyncServiceDocumentModel>();
            moqSyncServiceDocumentModel.SetupGet(x => x.Configuration).Returns(moqConfiguration.Object);
            var moqWordRibbon = new Mock<IWordRibbon>();
            var moqGetLinkTypesResult = new Mock<ICollection<ITFSWorkItemLinkType>>();

            // Act
            var getWorkItemsPanelViewModel = new GetWorkItemsPanelViewModel(moqtfsService.Object, moqSyncServiceDocumentModel.Object, moqWordRibbon.Object);

            var allWorkItems = new Dictionary<int,string>();
            var filteredList = new List<int>();
            allWorkItems.Add(1,"Bug");
            allWorkItems.Add(2, "Requirement");
            filteredList.Add(1);
            var amountOfFilteredWorkItems = filteredList.Count;
            var _idsOfFilteredWorkItems = filteredList;

            var result = getWorkItemsPanelViewModel.CompareCountsOfWorkItems(allWorkItems, amountOfFilteredWorkItems, _idsOfFilteredWorkItems);
            var expectedResult = new SortedDictionary<string, int>();
            expectedResult.Add("Requirement", 1);
            var firstObjectFromExpectedResult = expectedResult.First();
            var firstObjectFromRealResult = result.First();

            // Assert
            Assert.AreEqual(firstObjectFromExpectedResult.Key, firstObjectFromRealResult.Key);
            Assert.AreEqual(firstObjectFromExpectedResult.Value, firstObjectFromRealResult.Value);


        }
        #endregion
    }
}
