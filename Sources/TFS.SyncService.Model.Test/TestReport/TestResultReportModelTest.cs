#region Usings
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using AIT.TFS.SyncService.Contracts.Model;
using Moq;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.Contracts.TestCenter;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Model.TestReport;
using AIT.TFS.SyncService.Contracts.BuildCenter;
#endregion

namespace TFS.SyncService.Model.Test.Unit.TestReport
{
    [TestClass]
    public class TestResultReportModelTest
    {
        #region TestMethods
        [TestMethod]
        public void TestResultReportModelCreatingViaCmdGivesNoExeptions()
        {
            //new Sync
            var mySyncServiceDocModelMock = GetISyncServiceMockWithDefault();
            var myTfsTestAdapterMock = new Mock<ITfsTestAdapter>();
            var myTfsTestPlan = new Mock<ITfsTestPlan>();
            myTfsTestPlan.SetupGet(g => g.Name).Returns("Plan1");

            var tempList = new List<ITfsTestPlan> { myTfsTestPlan.Object };
            myTfsTestAdapterMock.Setup(g => g.AvailableTestPlans).Returns(tempList);
            var myTfsServerBuild = new Mock<ITfsServerBuild>();
            myTfsServerBuild.SetupGet(g => g.Name).Returns("All");
            myTfsServerBuild.SetupGet(g => g.BuildNumber).Returns(string.Empty);
            myTfsTestAdapterMock.Setup(g => g.GetAvailableServerBuilds()).Returns(new List<ITfsServerBuild> { myTfsServerBuild.Object });
            myTfsTestAdapterMock.Setup(x => x.GetTestConfigurationsForTestPlanWithTestResults(myTfsServerBuild.Object, myTfsTestPlan.Object)).Returns(new List<ITfsTestConfiguration>());
            var myProgressCancellationMock = new Mock<ITestReportingProgressCancellationService>();

            var testReportModelMock = new TestResultReportModel(mySyncServiceDocModelMock.Object, myTfsTestAdapterMock.Object, myTfsTestPlan.Object, null, false, 0, 
                DocumentStructureType.TestPlanHierarchy, TestCaseSortType.IterationPath, false, ConfigurationPositionType.AboveTestPlan, false, false, string.Empty, null, myProgressCancellationMock.Object);

            Assert.IsNotNull(testReportModelMock);
            Assert.IsNotNull(testReportModelMock.SelectedTestPlan);

        }
        #endregion

        private Mock<ISyncServiceDocumentModel> GetISyncServiceMockWithDefault()
        {
            var mySyncServiceDocModelMock = new Mock<ISyncServiceDocumentModel>();
            var myConfiguration = new Mock<IConfiguration>();
            var myConfigurationTest = new Mock<IConfigurationTest>();
            var myConfigurationTestResult = new Mock<IConfigurationTestResult>();
            var myTestResultRepDefaultVal = new Mock<ITestResultReportDefault>();
            myTestResultRepDefaultVal.SetupGet(g => g.SelectBuild).Returns("All");
            myTestResultRepDefaultVal.SetupGet(g => g.SelectTestPlan).Returns(string.Empty);
            myTestResultRepDefaultVal.SetupGet(g => g.SelectTestConfiguration).Returns(string.Empty);
            myTestResultRepDefaultVal.SetupGet(g => g.SkipLevels).Returns(0);
            myTestResultRepDefaultVal.SetupGet(g => g.CreateDocumentStructure).Returns(true);
            myTestResultRepDefaultVal.SetupGet(g => g.DocumentStructureType).Returns(DocumentStructureType.TestPlanHierarchy);
            myTestResultRepDefaultVal.SetupGet(g => g.IncludeTestConfigurations).Returns(false);
            myTestResultRepDefaultVal.SetupGet(g => g.ConfigurationPositionType).Returns(ConfigurationPositionType.AboveTestPlan);
            myTestResultRepDefaultVal.SetupGet(g => g.TestCaseSortType).Returns(TestCaseSortType.None);
            myTestResultRepDefaultVal.SetupGet(g => g.IncludeMostRecentTestResult).Returns(false);
            myTestResultRepDefaultVal.SetupGet(g => g.IncludeMostRecentTestResultForAllSelectedConfigurations).Returns(false);

            myConfigurationTestResult.SetupGet(g => g.DefaultValues).Returns(myTestResultRepDefaultVal.Object);
            myConfigurationTest.SetupGet(g => g.ConfigurationTestResult).Returns(myConfigurationTestResult.Object);
            myConfiguration.SetupGet(g => g.ConfigurationTest).Returns(myConfigurationTest.Object);
            mySyncServiceDocModelMock.SetupGet(s => s.Configuration).Returns(myConfiguration.Object);

            return mySyncServiceDocModelMock;
        }
    }
}
