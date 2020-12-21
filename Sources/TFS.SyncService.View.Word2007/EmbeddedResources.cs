namespace TFS.SyncService.View.Word2007
{
    #region Usings
    using System;
    using System.Globalization;
    using System.IO;
    using System.Reflection;
    using AIT.TFS.SyncService.Contracts;
    using AIT.TFS.SyncService.Factory;
    #endregion

    /// <summary>
    /// Exports embedded resources to local folder
    /// </summary>
    public static class EmbeddedResources
    {
        /// <summary>
        /// Gets the root path for exported resources
        /// </summary>
        private static string LocalAppData
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                    Constants.ApplicationCompany, Constants.ApplicationName);
            }
        }

        /// <summary>
        /// Export all resources.
        /// </summary>
        public static void ExportResources()
        {
            ExportResource("Templates", "MSFForAgile(2010).UserStory.xml");
            ExportResource("Templates", "MSFForAgile(2010).Issue.xml");
            ExportResource("Templates", "MSFForAgile(2010).Task.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Requirement.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Task.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Risk.xml");
            ExportResource("Templates", "Legacy(2010).Requirement.xml");
            ExportResource("Templates", "Legacy(2010).Task.xml");
            ExportResource("Templates", "Legacy(2010).ChangeRequest.xml");
            ExportResource("Templates", "Legacy(2010).Bug.xml");
            ExportResource("Templates", "Legacy(2010).Issue.xml");
            ExportResource("Templates", "Legacy(2010).Review.xml");
            ExportResource("Templates", "Legacy(2010).Risk.xml");
            ExportResource("Templates", "Legacy(2010).TestCase.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Header.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Bug.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.ProductBacklogItem.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Common.TestPlan.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Common.TestSuite.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Link.ActionSimpleList.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Link.AffectedByHeader.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Link.AffectsHeader.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Link.LinkSimpleList.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Link.AttachmentSimpleList.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Link.TestConfigurationHeader.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Link.TestConfigurationRow.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Result.TestCase.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Result.TestConfigurationHeader.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Result.TestConfigurationRow.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Result.TestResultHeader.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Result.TestResultRow.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Specification.TestCaseHeader.xml");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.Test.Specification.TestCaseRow.xml");

            ExportResource("Templates", "MSFForAgile(2010).w2t");
            ExportResource("Templates", "MSFForAgile(2010)DE.w2t");
            ExportResource("Templates", "MSFForAgile(2012).w2t");
            ExportResource("Templates", "MSFForAgile(2013).w2t");
            ExportResource("Templates", "MSFForCMMI(2010).w2t");
            ExportResource("Templates", "MSFForCMMI(2010)DE.w2t");
            ExportResource("Templates", "MSFForCMMI(2012).w2t");
            ExportResource("Templates", "MSFForCMMI(2013).w2t");
            ExportResource("Templates", "MSFForCMMI.Test.Common.TestPlan.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Common.TestSuite.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Link.ActionSimpleList.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Link.AffectedByHeader.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Link.AffectsHeader.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Link.LinkSimpleList.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Link.AttachmentSimpleList.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Link.TestConfigurationHeader.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Link.TestConfigurationRow.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Result.TestCase.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Result.TestConfigurationHeader.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Result.TestConfigurationRow.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Result.TestResultHeader.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Result.TestResultRow.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Specification.TestCaseHeader.xml");
            ExportResource("Templates", "MSFForCMMI.Test.Specification.TestCaseRow.xml");

            ExportResource("Templates", "Legacy(2010).w2t");
            ExportResource("Templates", "MicrosoftVisualStudioScrum.w2t");
            ExportResource("Templates", "MicrosoftVisualStudioScrum(2012).w2t");
            ExportResource("Templates", "MicrosoftVisualStudioScrum(2013).w2t");

            ExportResource("Templates", "TFSFeature(2013).xml");

            ExportResource("Templates", "MSFForCMMIHierarchyExample(2013).w2t");
            ExportResource("Templates", "MSFForCMMI(2010).Feature_Level0.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Feature_Level1.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Feature_Level2.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Feature_Level3.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Feature_Level4.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Feature_Level5.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Requirement_Level0.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Requirement_Level1.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Requirement_Level2.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Requirement_Level3.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Requirement_Level4.xml");
            ExportResource("Templates", "MSFForCMMI(2010).Requirement_Level5.xml");

            ExportResource("Templates", "Agile.UserStory(2015).xml");
            ExportResource("Templates", "Feature(2015).xml");
            ExportResource("Templates", "Epic(2015).xml");
            ExportResource("Templates", "Scrum.PBI(2015).xml");
            ExportResource("Templates", "Bug(2015).xml");
            ExportResource("Templates", "CMMI.Requirement(2015).xml");
            ExportResource("Templates", "CMMI.ChangeRequest(2019).xml");            
            ExportResource("Templates", "Test.Common.TestPlan.xml");
            ExportResource("Templates", "Test.Common.TestSuite.xml");
            ExportResource("Templates", "Test.Link.ActionList.xml");
            ExportResource("Templates", "Test.Link.AffectedByHeader.xml");
            ExportResource("Templates", "Test.Link.AffectsHeader.xml");
            ExportResource("Templates", "Test.Link.AttachmentSimpleList.xml");
            ExportResource("Templates", "Test.Link.IterationParameterList.xml");
            ExportResource("Templates", "Test.Link.LinkSimpleList.xml");
            ExportResource("Templates", "Test.Link.TestConfigurationHeader.xml");
            ExportResource("Templates", "Test.Link.TestConfigurationRow.xml");
            ExportResource("Templates", "Test.Result.TestCase.xml");
            ExportResource("Templates", "Test.Result.TestConfigurationHeader.xml");
            ExportResource("Templates", "Test.Result.TestConfigurationRow.xml");
            ExportResource("Templates", "Test.Result.TestResultHeader.xml");
            ExportResource("Templates", "Test.Result.TestResultRow.xml");
            ExportResource("Templates", "Test.Specification.TestCaseHeader.xml");
            ExportResource("Templates", "UnorderedList.xml");
            ExportResource("Templates", "Test.Specification.TestCaseRow.xml");
            ExportResource("Templates", "Scrum(2015).w2t");
            ExportResource("Templates", "Agile(2015).w2t");
            ExportResource("Templates", "CMMI(2015).w2t");
            ExportResource("Templates", "CMMI(2019).w2t");
            ExportResource("Templates", "Basic(2019).w2t");
            ExportResource("Templates", "Basic.Issue(2019).xml");
            ExportResource("Templates", "Epic(2019).xml");

            ExportResource("Templates", "W2T.xsd");

            ExportResource("Resources", "standard.png");
            ExportResource("Help", "AIT WordToTFS User Guide.pdf");
        }

        /// <summary>
        /// Exports a resource from the namespace and copy
        /// it to the local app data folder.
        /// "TFS.SyncService.View.Word2007.Help.WordToTFS.PDF -> AppData\AIT\Word2TFS\Help\WordToTFS.PDF
        /// </summary>
        /// <param name="subFolder">Subfolder of the resource to export.</param>
        /// <param name="fileName">Filename of the resource to export</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "We need to catch all exceptions and log all of them.")]
        private static void ExportResource(string subFolder, string fileName)
        {
            try
            {
                var path = Path.Combine(LocalAppData, subFolder);
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);

                var source = $"TFS.SyncService.View.Word2007.{subFolder}.{fileName}";
                var target = Path.Combine(path, fileName);

                using (
                    Stream readStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(source),
                           writeStream = new FileStream(target, FileMode.Create, FileAccess.Write))
                {
                    if (readStream != null) readStream.CopyTo(writeStream);
                }
            }

            catch (Exception e)
            {
                SyncServiceTrace.LogException(e);
            }

        }
    }
}