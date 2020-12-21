#region Usings
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.ProgressService;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.Helper;
using System;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Service.InfoStorage;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
#endregion
// ReSharper disable InconsistentNaming

namespace TFS.SyncService.Model.Test.Unit.Help
{
    [TestClass]
    public class TestReportingProgressCancellationServiceTest
    {
        #region TestMethods
        [TestMethod]
        public void TestReportingProgressCancellationServiceTest_ContinueFalseWhenProgressCanceledTrue()
        {
            //Arrange
            var progressService = new Mock<IProgressService>();
            progressService.SetupGet(x => x.ProgressCanceled).Returns(true);
            SyncServiceFactory.RegisterService(progressService.Object);
            var infoStorageService = new Mock<IInfoStorageService>();
            SyncServiceFactory.RegisterService(infoStorageService.Object);
            var cancelTrueServiceInstance = new TestReportingProgressCancellationService(true);

            //Act
            var shouldContinue = cancelTrueServiceInstance.CheckIfContinue();
            Assert.AreEqual(shouldContinue, false);
        }

        [TestMethod]
        public void TestReportingProgressCancellationServiceTest_ContinueFalseWhenProgressNotCanceledButErrorsExist()
        {
            var progressService = new Mock<IProgressService>();
            progressService.SetupGet(x => x.ProgressCanceled).Returns(false);
            SyncServiceFactory.RegisterService(progressService.Object);

            var infoStorageService = new Mock<IInfoStorageService>();
            infoStorageService.SetupGet(x => x.UserInformation).
                Returns(new InfoCollection<IUserInformation>() { new UserInformation() { Type = UserInformationType.Error } });

            SyncServiceFactory.RegisterService(infoStorageService.Object);
            var cancelTrueServiceInstance = new TestReportingProgressCancellationService(true);


            var shouldContinue = cancelTrueServiceInstance.CheckIfContinue();
            Assert.AreEqual(shouldContinue, false);
        }

        [TestMethod]
        public void TestReportingProgressCancellationServiceTest_ContinueFalseWhenProgressNotCanceledAndNoeErrorExist()
        {
            var progressService = new Mock<IProgressService>();
            progressService.SetupGet(x => x.ProgressCanceled).Returns(false);
            SyncServiceFactory.RegisterService(progressService.Object);

            var infoStorageService = new Mock<IInfoStorageService>();
            infoStorageService.SetupGet(x => x.UserInformation).
                Returns(new InfoCollection<IUserInformation>() { new UserInformation() { Type = UserInformationType.Warning } });

            SyncServiceFactory.RegisterService(infoStorageService.Object);
            var cancelTrueServiceInstance = new TestReportingProgressCancellationService(true);

            var shouldContinue = cancelTrueServiceInstance.CheckIfContinue();
            Assert.AreEqual(shouldContinue, true);
        }
        #endregion
    }
}
