#region Usings
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Configuration.TestReport;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Model.TestReport;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
#endregion

namespace TFS.SyncService.Model.Test.Unit
{
    /// <summary>
    /// Tests the result report model helper
    /// </summary>
    [TestClass]
    public class TestReportHelperTest
    {
        #region TestMehods
        /// <summary>
        /// Should set hyperlink base if configured.
        /// </summary>
        [TestMethod]
        public void TestReportHelper_EvaluateAttachmentLink_ShouldSetHyperlinkBase()
        {
            // setup
            var folder = Path.GetTempPath();

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupAllProperties();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");

            var attachment = new Mock<ITestAttachment>();
            attachment.SetupGet(f => f.Length).Returns(1);
            attachment.SetupGet(f => f.IsComplete).Returns(true);
            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("LocalPath", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("ThisShouldNotBeReturnedAsThisPropertyIsFake");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("LocalPath");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.DownloadAndLinkToLocalFile);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(x => x.ConfigurationTest.SetHyperlinkBase).Returns(true);
            configuration.SetupGet(x => x.ConfigurationTest.ShowHyperlinkBaseMessageBoxes).Returns(false);
            
            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify
            wordTestAdapter.VerifySet(x => x.HyperlinkBase = folder);
        }

        /// <summary>
        /// Should set hyperlink base if configured.
        /// </summary>
        [TestMethod]
        public void TestReportHelper_EvaluateObjectQueriesWithWorkItemName_ShouldInnovacateAddCustomMethod()
        {
            // setup
            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupAllProperties();
            var testResultDetail = new Mock<ITfsTestResultDetail>();

            var destElement = new Mock<IElement>();
            destElement.SetupGet(x => x.ItemName).Returns("Change Request");

            var objectQuery = new Mock<IObjectQuery>();
            objectQuery.SetupGet(x => x.DestinationElements).Returns(new List<IElement> { destElement.Object });
            objectQuery.SetupGet(x => x.Name).Returns("TestObjectQuery");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("TestObjectQuery");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ObjectQueries).Returns(new List<IObjectQuery> { objectQuery.Object });

            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(x => x.ConfigurationTest.SetHyperlinkBase).Returns(true);
            configuration.SetupGet(x => x.ConfigurationTest.ShowHyperlinkBaseMessageBoxes).Returns(false);


            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns("name");
            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns(() => null);


            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);
            
            // verify that the AddCustomObject method was innovated
            testResultDetail.Verify();
            testResultDetail.Verify(x => x.AddCustomObjects("TestObjectQuery", It.IsAny<object>()));
        }

        /// <summary>
        /// Should set hyperlink base if configured.
        /// </summary>
        [TestMethod]
        public void TestReportHelper_EvaluateObjectQueriesWithBuilds_ShouldInnovacateAddCustomMethod()
        {
            // setup
            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupAllProperties();
            var testResultDetail = new Mock<ITfsTestResultDetail>();

            var destElement = new Mock<IElement>();
            destElement.SetupGet(x => x.ItemName).Returns("Builds");

            var objectQuery = new Mock<IObjectQuery>();
            objectQuery.SetupGet(x => x.DestinationElements).Returns(new List<IElement> { destElement.Object });
            objectQuery.SetupGet(x => x.Name).Returns("TestObjectQuery");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("TestObjectQuery");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ObjectQueries).Returns(new List<IObjectQuery> { objectQuery.Object });

            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(x => x.ConfigurationTest.SetHyperlinkBase).Returns(true);
            configuration.SetupGet(x => x.ConfigurationTest.ShowHyperlinkBaseMessageBoxes).Returns(false);


            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns("name");
            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns(() => null);


            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify that the AddCustomObject method was innovated
            testResultDetail.Verify();
            testResultDetail.Verify(x => x.AddCustomObjects("TestObjectQuery", It.IsAny<object>()));
        }

        /// <summary>
        /// Should set hyperlink base if configured.
        /// </summary>
        [TestMethod]
        public void TestReportHelper_EvaluateObjectQueriesWithConfigurations_ShouldInnovacateAddCustomMethod()
        {
            // setup
            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupAllProperties();
            var testResultDetail = new Mock<ITfsTestResultDetail>();

            var destElement = new Mock<IElement>();
            destElement.SetupGet(x => x.ItemName).Returns("Configurations");

            var objectQuery = new Mock<IObjectQuery>();
            objectQuery.SetupGet(x => x.DestinationElements).Returns(new List<IElement> { destElement.Object });
            objectQuery.SetupGet(x => x.Name).Returns("TestObjectQuery");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("TestObjectQuery");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ObjectQueries).Returns(new List<IObjectQuery> { objectQuery.Object });

            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(x => x.ConfigurationTest.SetHyperlinkBase).Returns(true);
            configuration.SetupGet(x => x.ConfigurationTest.ShowHyperlinkBaseMessageBoxes).Returns(false);


            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns("name");
            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns(() => null);

            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify that the AddCustomObject method was innovated
            testResultDetail.Verify();
            testResultDetail.Verify(x => x.AddCustomObjects("TestObjectQuery", It.IsAny<object>()));
            testResultDetail.Setup(x => x.GetCustomObject("TestObjectQuery")).Returns(It.IsNotNull<object>);

        }

        /// <summary>
        /// Should set hyperlink base if configured.
        /// </summary>
        [TestMethod]
        public void TestReportHelper_EvaluateObjectQueriesWithTestResults_ShouldInnovacateAddCustomMethod()
        {
            // setup
            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupAllProperties();
            var testResultDetail = new Mock<ITfsTestResultDetail>();

            var destElement = new Mock<IElement>();
            destElement.SetupGet(x => x.ItemName).Returns("Test Results");

            var objectQuery = new Mock<IObjectQuery>();
            objectQuery.SetupGet(x => x.DestinationElements).Returns(new List<IElement> { destElement.Object });
            objectQuery.SetupGet(x => x.Name).Returns("TestObjectQuery");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("TestObjectQuery");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ObjectQueries).Returns(new List<IObjectQuery> { objectQuery.Object });

            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(x => x.ConfigurationTest.SetHyperlinkBase).Returns(true);
            configuration.SetupGet(x => x.ConfigurationTest.ShowHyperlinkBaseMessageBoxes).Returns(false);


            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns("name");
            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns(() => null);


            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify that the AddCustomObject method was innovated
            testResultDetail.Verify();
            testResultDetail.Verify(x => x.AddCustomObjects("TestObjectQuery", It.IsAny<object>() ));

           }

        /// <summary>
        /// Should set hyperlink base if configured.
        /// </summary>
        [TestMethod]
        public void TestReportHelper_EvaluateObjectQueriesWithTestCases_ShouldInnovacateAddCustomMethod()
        {
            // setup
            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupAllProperties();
            var testResultDetail = new Mock<ITfsTestResultDetail>();

            var destElement = new Mock<IElement>();
            destElement.SetupGet(x => x.ItemName).Returns("Test Cases");

            var objectQuery = new Mock<IObjectQuery>();
            objectQuery.SetupGet(x => x.DestinationElements).Returns(new List<IElement> { destElement.Object });
            objectQuery.SetupGet(x => x.Name).Returns("TestObjectQuery");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("TestObjectQuery");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ObjectQueries).Returns(new List<IObjectQuery> { objectQuery.Object });

            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(x => x.ConfigurationTest.SetHyperlinkBase).Returns(true);
            configuration.SetupGet(x => x.ConfigurationTest.ShowHyperlinkBaseMessageBoxes).Returns(false);


            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns("name");
            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns(() => null);


            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify that the AddCustomObject method was innovated
            testResultDetail.Verify();
            testResultDetail.Verify(x => x.AddCustomObjects("TestObjectQuery", null));

        }

        /// <summary>
        /// Should set hyperlink base if configured.
        /// </summary>
        [TestMethod]
        public void TestReportHelper_EvaluateObjectQueriesWithUnknownName_ShouldNotInnovacateAddCustomMethod()
        {



            // setup
            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupAllProperties();
            var testResultDetail = new Mock<ITfsTestResultDetail>();

            var destElement = new Mock<IElement>();
            destElement.SetupGet(x => x.ItemName).Returns("Test Cases");

            var objectQuery = new Mock<IObjectQuery>();
            objectQuery.SetupGet(x => x.DestinationElements).Returns(new List<IElement> { destElement.Object });
            objectQuery.SetupGet(x => x.Name).Returns("TestObjectQuery");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("TestObjectQuery");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ObjectQueries).Returns(new List<IObjectQuery> { objectQuery.Object });

            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(x => x.ConfigurationTest.SetHyperlinkBase).Returns(true);
            configuration.SetupGet(x => x.ConfigurationTest.ShowHyperlinkBaseMessageBoxes).Returns(false);

            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns("name");
            testResultDetail.Setup(x => x.GetCustomObject("Name")).Returns(() => null);

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);
          
            // verify that the AddCustomObject method was innovated
            testResultDetail.Verify(x => x.AddCustomObjects("TestObjectQuery", It.IsAny<object>()));
     
        }

        /// <summary>
        /// Tests standard behaviour for links
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        public void TestReportHelper_EvaluateAttachmentLink_ShouldDownloadAndReplaceBookmarkWithPropertyValueAndHyperlink()
        {
            // setup
            var folder = Path.GetTempPath();
            var localPath = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");

            // ReSharper disable once ImplicitlyCapturedClosure
            wordTestAdapter.Setup(x => x.ReplaceBookmarkHyperlink(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Callback((string x, string y, string z) => localPath = z);

            var attachment = new Mock<ITestAttachment>();

            attachment.SetupGet(f => f.Length).Returns(1);
            attachment.SetupGet(f => f.IsComplete).Returns(true);

            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("Comment", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyComment");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("Comment");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.DownloadAndLinkToLocalFile);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify
            Assert.IsTrue(localPath.EndsWith("MyAttachmentName", StringComparison.OrdinalIgnoreCase));
            attachment.Verify(x => x.DownloadToFile(folder + localPath), Times.Once());
            wordTestAdapter.Verify(x => x.ReplaceBookmarkHyperlink("Bookmark", "MyComment", It.IsAny<string>()), Times.Once());
            wordTestAdapter.VerifySet(x => x.HyperlinkBase = It.IsAny<string>(), Times.Never());
        }

        /// <summary>
        /// Tests standard behaviour for links
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        public void TestReportHelper_EvaluateAttachmentLink_WithoutGuid_ShouldDownloadAndReplaceBookmarkWithPropertyValueAndHyperlink()
        {
            // setup
            var folder = Path.GetTempPath();
            var localPath = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");

            // ReSharper disable once ImplicitlyCapturedClosure
            wordTestAdapter.Setup(x => x.ReplaceBookmarkHyperlink(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Callback((string x, string y, string z) => localPath = z);

            var attachment = new Mock<ITestAttachment>();

            attachment.SetupGet(f => f.Length).Returns(1);
            attachment.SetupGet(f => f.IsComplete).Returns(true);

            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("Comment", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyComment");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("Comment");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.DownloadAndLinkToLocalFile);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(e => e.AttachmentFolderMode).Returns(AttachmentFolderMode.WithoutGuid);

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify
            Assert.IsTrue(localPath.EndsWith("MyAttachmentName", StringComparison.OrdinalIgnoreCase));
            attachment.Verify(x => x.DownloadToFile(folder + localPath), Times.Once());
            wordTestAdapter.Verify(x => x.ReplaceBookmarkHyperlink("Bookmark", "MyComment", It.IsAny<string>()), Times.Once());
            wordTestAdapter.VerifySet(x => x.HyperlinkBase = It.IsAny<string>(), Times.Never());
        }

        /// <summary>
        /// Tests downloading attachments without hyperlinks
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        public void TestReportHelper_EvaluateAttachmentLink_ShouldDownloadAndReplaceBookmarkWithPropertyValueWithoutHyperlink()
        {
            // setup
            var folder = Path.GetTempPath();
            // ReSharper disable once NotAccessedVariable
            var localPath = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");

            var attachment = new Mock<ITestAttachment>();
            attachment.SetupGet(f => f.Length).Returns(1);
            attachment.SetupGet(f => f.IsComplete).Returns(true);

            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("Comment", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyComment");
            attachment.Setup(x => x.DownloadToFile(It.IsAny<string>())).Callback((string x) => localPath = x);


            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("Comment");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.DownloadOnly);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify
            attachment.Verify(x => x.DownloadToFile(It.IsRegex(folder.Replace("\\", "\\\\") + ".*" + "MyAttachmentName")), Times.Once());
            wordTestAdapter.Verify(x => x.ReplaceBookmarkText("Bookmark", "MyComment", PropertyValueFormat.PlainText, null), Times.Once());
            wordTestAdapter.Verify(x => x.ReplaceBookmarkHyperlink(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>()), Times.Never());
            wordTestAdapter.VerifySet(x => x.HyperlinkBase = It.IsAny<string>(), Times.Never());
        }

        /// <summary>
        /// Tests server links for attachments links
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        public void TestReportHelper_EvaluateAttachmentLink_ShouldLinkToServer()
        {
            // setup
            //var folder = Path.GetTempPath();
            //var link = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns("MyFolder");
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");
            //wordTestAdapter.Setup(x => x.ReplaceBookmarkHyperlink(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Callback((string x, string y, string z) => link = z);

            var attachment = new Mock<ITestAttachment>();
            attachment.SetupGet(f => f.Length).Returns(1);
            attachment.SetupGet(f => f.IsComplete).Returns(true);
            attachment.Setup(x => x.Uri).Returns(new Uri("http://wwww.test.de"));
            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("Comment", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyComment");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("Comment");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.LinkToServerVersion);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify
            attachment.Verify(x => x.DownloadToFile(It.IsAny<string>()), Times.Never());
            wordTestAdapter.Verify(x => x.ReplaceBookmarkHyperlink("Bookmark", "MyComment", (new Uri("http://wwww.test.de").ToString())), Times.Once());
            wordTestAdapter.VerifySet(x => x.HyperlinkBase = It.IsAny<string>(), Times.Never());
        }

        /// <summary>
        /// Attachment links have access to a fake "LocalPath" property that should return local path
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        public void TestReportHelper_EvaluateAttachmentLink_ShouldSimulateLocalPathProperty()
        {
            // setup
            var folder = Path.GetTempPath();
            var localPath = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");
            wordTestAdapter.Setup(x => x.ReplaceBookmarkHyperlink(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Callback((string x, string y, string z) => localPath = z);

            var attachment = new Mock<ITestAttachment>();
            attachment.SetupGet(f => f.Length).Returns(1);
            attachment.SetupGet(f => f.IsComplete).Returns(true);
            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("LocalPath", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("ThisShouldNotBeReturnedAsThisPropertyIsFake");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("LocalPath");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.DownloadAndLinkToLocalFile);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify
            wordTestAdapter.Verify(x => x.ReplaceBookmarkHyperlink("Bookmark", localPath, localPath), Times.Once());
        }

        /// <summary>
        /// Attachment links have access to a fake "LocalPath" property that should return local path
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        public void TestReportHelper_EvaluateAttachmentLink_ShouldSimulateLocalFileNameProperty()
        {
            // setup
            var folder = Path.GetTempPath();
            var localPath = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");
            wordTestAdapter.Setup(x => x.ReplaceBookmarkHyperlink(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Callback((string x, string y, string z) => localPath = z);

            var attachment = new Mock<ITestAttachment>();
            attachment.SetupGet(f => f.Length).Returns(1);
            attachment.SetupGet(f => f.IsComplete).Returns(true);
            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("LocalPath", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("ThisShouldNotBeReturnedAsThisPropertyIsFake");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("LocalFilename");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.DownloadAndLinkToLocalFile);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify
            var localFilename = localPath.Replace(wordTestAdapter.Object.AttachmentFolder + Path.DirectorySeparatorChar, "");
            wordTestAdapter.Verify(x => x.ReplaceBookmarkHyperlink("Bookmark", localFilename, localPath), Times.Once());
        }

        /// <summary>
        /// Tests standard behaviour for links
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestReportHelper_EvaluateAttachmentLink_CompleteAttachmentWithLength0_ShouldFail()
        {
            // setup
            var folder = Path.GetTempPath();
            // ReSharper disable once NotAccessedVariable
            var localPath = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");

            wordTestAdapter.Setup(x => x.ReplaceBookmarkHyperlink(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Callback((string x, string y, string z) => localPath = z);

            var attachment = new Mock<ITestAttachment>();

            attachment.SetupGet(f => f.Length).Returns(0);
            attachment.SetupGet(f => f.IsComplete).Returns(true);

            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("Comment", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyComment");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("Comment");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.DownloadAndLinkToLocalFile);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(e => e.AttachmentFolderMode).Returns(AttachmentFolderMode.WithoutGuid);

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

        }

        /// <summary>
        /// Tests standard behaviour for links
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestReportHelper_EvaluateAttachmentLink_IncompleteAttachment_ShouldFail()
        {
            // setup
            var folder = Path.GetTempPath();
            // ReSharper disable once NotAccessedVariable
            var localPath = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");

            wordTestAdapter.Setup(x => x.ReplaceBookmarkHyperlink(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Callback((string x, string y, string z) => localPath = z);

            var attachment = new Mock<ITestAttachment>();

            attachment.SetupGet(f => f.Length).Returns(1);
            attachment.SetupGet(f => f.IsComplete).Returns(false);

            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("Comment", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyComment");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("Comment");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.DownloadAndLinkToLocalFile);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(e => e.AttachmentFolderMode).Returns(AttachmentFolderMode.WithoutGuid);

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

        }

        /// <summary>
        /// Tests standard behaviour for links
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestReportHelper_EvaluateAttachmentLink_IncompleteAttachmentWithLength0_ShouldFail()
        {
            // setup
            var folder = Path.GetTempPath();
            // ReSharper disable once NotAccessedVariable
            var localPath = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");

            wordTestAdapter.Setup(x => x.ReplaceBookmarkHyperlink(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Callback((string x, string y, string z) => localPath = z);

            var attachment = new Mock<ITestAttachment>();

            attachment.SetupGet(f => f.Length).Returns(0);
            attachment.SetupGet(f => f.IsComplete).Returns(false);

            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("Comment", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyComment");

            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("Comment");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.AttachmentLink.Mode).Returns(AttachmentLinkMode.DownloadAndLinkToLocalFile);

            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });
            configuration.SetupGet(e => e.AttachmentFolderMode).Returns(AttachmentFolderMode.WithoutGuid);

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

        }

        /// <summary>
        /// Test the insertion of custom bookmarks
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"), TestMethod]
        public void TestReportHelper_EvaluateOwnBookmarks_ShouldInsertOwnBookmark()
        {
            // setup
            var folder = Path.GetTempPath();
            // ReSharper disable once NotAccessedVariable
            var localPath = string.Empty;

            var tfsTestAdapter = new Mock<ITfsTestAdapter>();
            var wordTestAdapter = new Mock<IWord2007TestReportAdapter>();
            wordTestAdapter.SetupGet(x => x.DocumentPath).Returns(folder);
            wordTestAdapter.SetupGet(x => x.AttachmentFolder).Returns("Attachments");

            var attachment = new Mock<ITestAttachment>();
            attachment.SetupGet(f => f.Length).Returns(1);
            attachment.SetupGet(f => f.IsComplete).Returns(true);
            var testResultDetail = new Mock<ITfsTestResultDetail>();
            testResultDetail.SetupGet(x => x.AssociatedObject).Returns(attachment.Object);
            testResultDetail.Setup(x => x.PropertyValue("Name", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyAttachmentName");
            testResultDetail.Setup(x => x.PropertyValue("Comment", It.IsAny<Func<IEnumerable, IEnumerable>>())).Returns("MyComment");
            attachment.Setup(x => x.DownloadToFile(It.IsAny<string>())).Callback((string x) => localPath = x);


            var replacement = new Mock<IConfigurationTestReplacement>();
            replacement.SetupGet(x => x.PropertyToEvaluate).Returns("Comment");
            replacement.SetupGet(x => x.Bookmark).Returns("Bookmark");
            replacement.SetupGet(x => x.WordBookmark).Returns("OwnBookmarkValue");


            var configuration = new Mock<IConfiguration> { DefaultValue = DefaultValue.Mock };
            configuration.Setup(x => x.ConfigurationTest.GetReplacements(It.IsAny<string>())).Returns(new List<IConfigurationTestReplacement> { replacement.Object });

            // test
            var helper = new TestReportHelper(tfsTestAdapter.Object, wordTestAdapter.Object, configuration.Object, () => false);
            helper.InsertTestResult("test", testResultDetail.Object);

            // verify
            wordTestAdapter.Verify(x => x.ReplaceBookmarkText("Bookmark", "MyComment", PropertyValueFormat.PlainText, "OwnBookmarkValue"), Times.Once());
            wordTestAdapter.VerifySet(x => x.HyperlinkBase = It.IsAny<string>(), Times.Never());
        }
        #endregion

    }
}
