#region Usings
using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using AIT.TFS.SyncService.Model.TestReport;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Contracts.Word;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Adapter.Word2007;
using AIT.TFS.SyncService.Model.Helper;
#endregion

namespace TFS.SyncService.Model.Test.Unit.TestReport
{
    [TestClass]
    public class TestSpecificationReportModelTest
    {
        #region TestMethods
        [TestMethod]
        public void TestReportModelTestDispatcherNullOnWordCallsNoExeptions()
        {
            //new Sync
            var mySyncServiceDocModelMock = GetISyncServiceMockWithDefault();
            var myTfsTestAdapterMock = new Mock<ITfsTestAdapter>();
            var myTfsTestPlan = new Mock<ITfsTestPlan>();
            myTfsTestPlan.SetupGet(g => g.Name).Returns("Plan1");

            var tempList = new List<ITfsTestPlan> { myTfsTestPlan.Object };
            myTfsTestAdapterMock.Setup(g => g.AvailableTestPlans).Returns(tempList);

            var myWordRibonMock = new Mock<IWordRibbon>();
            var myProgressCancellationMock = new Mock<ITestReportingProgressCancellationService>();

            var testReportModelMock = new TestSpecificationReportModel(mySyncServiceDocModelMock.Object, new ViewDispatcher(null), myTfsTestAdapterMock.Object, myWordRibonMock.Object, myProgressCancellationMock.Object);

            try
            {
                var myTestConfiguration = new Mock<ITfsTestConfiguration>();
                testReportModelMock.SelectedTestConfiguration = myTestConfiguration.Object;
                var plans = testReportModelMock.TestPlans;
                Assert.IsNull(plans);
                testReportModelMock.SelectedTestPlan = myTfsTestPlan.Object;
            }
            catch (NullReferenceException)
            {
                Assert.Fail("Dispatcher is null and must not be called.");
            }

        }

        [TestMethod]
        public void TestSpecReportModelCreatingViaCmdGivesNoExeptions()
        {
            //new Sync
            var mySyncServiceDocModelMock = GetISyncServiceMockWithDefault();
            var myTfsTestAdapterMock = new Mock<ITfsTestAdapter>();
            var myTfsTestPlan = new Mock<ITfsTestPlan>();
            myTfsTestPlan.SetupGet(g => g.Name).Returns("Plan1");

            var tempList = new List<ITfsTestPlan> { myTfsTestPlan.Object };
            myTfsTestAdapterMock.Setup(g => g.AvailableTestPlans).Returns(tempList);

            var myProgressCancellationMock = new Mock<ITestReportingProgressCancellationService>();


           var testReportModelMock = new TestSpecificationReportModel(mySyncServiceDocModelMock.Object, myTfsTestAdapterMock.Object, myTfsTestPlan.Object, null, true, 0, DocumentStructureType.TestPlanHierarchy, TestCaseSortType.IterationPath, myProgressCancellationMock.Object);

           Assert.IsNotNull(testReportModelMock);
           Assert.IsNotNull(testReportModelMock.SelectedTestPlan);

        }

        #endregion

        private Mock<ISyncServiceDocumentModel>  GetISyncServiceMockWithDefault()
        {
            var mySyncServiceDocModelMock = new Mock<ISyncServiceDocumentModel>();
            var myConfiguration = new Mock<IConfiguration>();
            var myConfigurationTest = new Mock<IConfigurationTest>();
            var myConfigurationTestSpec = new Mock<IConfigurationTestSpecification>();
            var myTestSpecRepDefaultVal = new Mock<ITestSpecReportDefault>();
            myTestSpecRepDefaultVal.SetupGet(g => g.SelectTestPlan).Returns(string.Empty);
            myTestSpecRepDefaultVal.SetupGet(g => g.SelectTestSuite).Returns(string.Empty);
            myTestSpecRepDefaultVal.SetupGet(g => g.CreateDocumentStructure).Returns(true);
            myTestSpecRepDefaultVal.SetupGet(g => g.DocumentStructureType).Returns(DocumentStructureType.TestPlanHierarchy);
            myTestSpecRepDefaultVal.SetupGet(g => g.IncludeTestConfigurations).Returns(false);
            myTestSpecRepDefaultVal.SetupGet(g => g.ConfigurationPositionType).Returns(ConfigurationPositionType.BeneathTestPlan);
            myTestSpecRepDefaultVal.SetupGet(g => g.TestCaseSortType).Returns(TestCaseSortType.None);

            myConfigurationTestSpec.SetupGet(g => g.DefaultValues).Returns(myTestSpecRepDefaultVal.Object);
            myConfigurationTest.SetupGet(g => g.ConfigurationTestSpecification).Returns(myConfigurationTestSpec.Object);
            myConfiguration.SetupGet(g => g.ConfigurationTest).Returns(myConfigurationTest.Object);
            mySyncServiceDocModelMock.SetupGet(s => s.Configuration).Returns(myConfiguration.Object);

            return mySyncServiceDocModelMock;
        }
    }
}
