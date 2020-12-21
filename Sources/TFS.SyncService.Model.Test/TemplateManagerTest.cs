#region Usings
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using AIT.TFS.SyncService.Contracts.TemplateManager;
using AIT.TFS.SyncService.Model.TemplateManagement;
using Microsoft.VisualStudio.TestTools.UnitTesting;
#endregion

namespace TFS.SyncService.Model.Test.Unit
{
    [TestClass]
    public class TemplateManagerTest
    {
        TemplateManager _templateManager = new TemplateManager();
        private string _testLocation;
        private Guid _guid = Guid.NewGuid();

        [TestInitialize]
        public void TestInitialize()
        {

            if (Directory.Exists(TemplateManager.RootCacheDirectory))
            {
                Directory.Delete(TemplateManager.RootCacheDirectory, true);
            }

            PrepareTestLocation();
            _templateManager = new TemplateManager();
        }
        private void PrepareTestLocation()
        {
            _testLocation = Path.Combine(Path.GetTempPath(), "WordToTfsUnitTestLocation");

            if (Directory.Exists(_testLocation))
            {
                foreach (var fileInfo in new DirectoryInfo(_testLocation).EnumerateFiles())
                {
                    fileInfo.Attributes = FileAttributes.Normal;
                }

                Directory.Delete(_testLocation, true);
            }

            Directory.CreateDirectory(_testLocation);
            while (!Directory.Exists(_testLocation))
            {
                Thread.Sleep(100);
            }

            // Export a simple test configuration with related files from resources
            var assembly = Assembly.GetExecutingAssembly();
            using (var fs = new FileStream(Path.Combine(_testLocation, "Legacy(2010).w2t"), FileMode.Create))
            {
                var manifestResourceStream = assembly.GetManifestResourceStream("TFS.SyncService.Model.Test.Unit.TestResources.Legacy(2010).w2t");
                Assert.IsNotNull(manifestResourceStream);
                manifestResourceStream.CopyTo(fs);

            }

            using (var fs = new FileStream(Path.Combine(_testLocation, "Legacy(2010).Requirement.xml"), FileMode.Create))
            {
                var manifestResourceStream = assembly.GetManifestResourceStream("TFS.SyncService.Model.Test.Unit.TestResources.Legacy(2010).Requirement.xml");
                Assert.IsNotNull(manifestResourceStream);

                manifestResourceStream.CopyTo(fs);

            }

            if (!Directory.Exists(TemplateManager.DefaultTemplateBundleLocation))
                Directory.CreateDirectory(TemplateManager.DefaultTemplateBundleLocation);
            // Export some more test configurations to the default template folder.
            using (var fs = new FileStream(Path.Combine(TemplateManager.DefaultTemplateBundleLocation, "Legacy(2010).w2t"), FileMode.Create))
            {
                var manifestResourceStream = assembly.GetManifestResourceStream("TFS.SyncService.Model.Test.Unit.TestResources.Legacy(2010).w2t");
                Assert.IsNotNull(manifestResourceStream);
                manifestResourceStream.CopyTo(fs);

            }

            using (var fs = new FileStream(Path.Combine(TemplateManager.DefaultTemplateBundleLocation, "Legacy(2010).Requirement.xml"), FileMode.Create))
            {
                var manifestResourceStream = assembly.GetManifestResourceStream("TFS.SyncService.Model.Test.Unit.TestResources.Legacy(2010).Requirement.xml");
                Assert.IsNotNull(manifestResourceStream);
                manifestResourceStream.CopyTo(fs);

            }
        }

        #region TestMethods
        /// <summary>
        /// Check if entry and cache entry are created and files actually copied.
        /// </summary>
        [TestMethod]
        public void TemplateManagerTest_AddTemplateBundle_ShouldAddTemplateToListShouldCreateCacheEntry()
        {
            _templateManager.AddTemplateBundle(new TemplateBundle("TestBundle", _testLocation, _guid));
            Assert.IsTrue(File.Exists(Path.Combine(TemplateManager.RootCacheDirectory, _guid.ToString(), "Legacy(2010).w2t")));

            // 2 because the default template also gets created.
            Assert.AreEqual(2, _templateManager.TemplateBundles.Count());
        }

        /// <summary>
        /// Check if entry and cache entry are removed.
        /// </summary>
        [TestMethod]
        public void TemplateManagerTest_RemoveTemplateBundle_ShouldRemoveTemplateToListShouldRemoveCacheEntry()
        {
            TemplateManagerTest_AddTemplateBundle_ShouldAddTemplateToListShouldCreateCacheEntry();
            _templateManager.RemoveTemplateBundle("TestBundle");

            Assert.IsFalse(Directory.Exists(Path.Combine(TemplateManager.RootCacheDirectory, _guid.ToString())));

            // 1 because the default template also was created.
            Assert.AreEqual(1, _templateManager.TemplateBundles.Count());
        }

        /// <summary>
        /// Load cached template if source is not available.
        /// </summary>
        [TestMethod]
        public void TemplateManagerTest_RemoveSourceLocation_ShouldLoadCachedTemplate()
        {
            // Add test location and remove it
            TemplateManagerTest_AddTemplateBundle_ShouldAddTemplateToListShouldCreateCacheEntry();
            Directory.Delete(_testLocation, true);

            var templateManager = new TemplateManager();

            Assert.AreEqual(2, templateManager.TemplateBundles.Count());
            Assert.IsTrue((templateManager.TemplateBundles.First(x => x.ShowName.Equals("TestBundle")).TemplateBundleState & TemplateBundleStates.TemplateBundleCached) != 0);
        }

        /// <summary>
        /// Deactivate the first template and see if the deactivation is stored and loaded correctly
        /// </summary>
        [TestMethod]
        public void TemplateManagerTest_DeactivateTemplate_ShouldSaveAndLoadDeactivation()
        {
            var firstName = _templateManager.AvailableTemplates.First().ShowName;
            _templateManager.AvailableTemplates.First(x => x.ShowName.Equals(firstName)).TemplateState = TemplateState.Disabled;

            var templateManager = new TemplateManager();
            Assert.IsTrue(templateManager.AvailableTemplates.First(x => x.ShowName.Equals(firstName)).TemplateState == TemplateState.Disabled);
        }

        /// <summary>
        /// Test whether subfolders are interpreted as server/TPC/Project restriction
        /// </summary>
        [TestMethod]
        public void TemplateManagerTest_UseProjectMappedFolderHierarchy_ShouldParseFolderStructureAndSaveSetting()
        {
            var path = Path.Combine(_testLocation, "TeamProjectCollectionUrl", "ProjectCollectionName", "TeamProjectName");
            Directory.CreateDirectory(path);
            foreach (var file in new DirectoryInfo(_testLocation).GetFiles("*.*", SearchOption.TopDirectoryOnly))
            {
                file.MoveTo(Path.Combine(path, file.Name));
            }

            var templateBundle = new TemplateBundle("ProjectMappedFolderHierarchy", _testLocation, _guid);
            templateBundle.HasProjectMappedFolderHierarchy = true;
            _templateManager.AddTemplateBundle(templateBundle);

            Assert.AreEqual("TeamProjectCollectionUrl", templateBundle.Templates.First().ServerName);
            Assert.AreEqual("ProjectCollectionName", templateBundle.Templates.First().ProjectCollectionName);
            Assert.AreEqual("TeamProjectName", templateBundle.Templates.First().ProjectName);

            // save, load and make sure setting was applied correctly
            _templateManager.SaveTemplateBundles();
            var templateManager = new TemplateManager();

            Assert.IsTrue(templateManager.TemplateBundles.First(x => x.ShowName.Equals("ProjectMappedFolderHierarchy")).HasProjectMappedFolderHierarchy);
        }

        /// <summary>
        /// Test whether the first folder hierarchy is ignored when using more then three subfolder levels
        /// </summary>
        [TestMethod]
        public void TemplateManagerTest_UseProjectMappedFolderHierarchy_LoadVariantFolder()
        {
            var path1 = Path.Combine(_testLocation, "Variant1", "TeamProjectCollectionUrl", "ProjectCollectionName", "TeamProjectName");
            var path2 = Path.Combine(_testLocation, "Variant2", "TeamProjectCollectionUrl", "ProjectCollectionName", "TeamProjectName");

            Directory.CreateDirectory(path1);
            Directory.CreateDirectory(path2);
            foreach (var file in new DirectoryInfo(_testLocation).GetFiles("*.*", SearchOption.TopDirectoryOnly))
            {
                file.CopyTo(Path.Combine(path1, file.Name));
                file.MoveTo(Path.Combine(path2, file.Name));
            }

            var templateBundle = new TemplateBundle("ProjectMappedFolderHierarchy", _testLocation, _guid);
            templateBundle.HasProjectMappedFolderHierarchy = true;
            _templateManager.AddTemplateBundle(templateBundle);

            Assert.AreEqual("TeamProjectCollectionUrl", templateBundle.Templates.First().ServerName);
            Assert.AreEqual("ProjectCollectionName", templateBundle.Templates.First().ProjectCollectionName);
            Assert.AreEqual("TeamProjectName", templateBundle.Templates.First().ProjectName);

            // save, load and make sure setting was applied correctly
            _templateManager.SaveTemplateBundles();
            var templateManager = new TemplateManager();

            // 3! Two variants and the default
            Assert.AreEqual(2, templateManager.TemplateBundles.First(x => x.ShowName.Equals("ProjectMappedFolderHierarchy")).Templates.Count);
            Assert.IsTrue(templateManager.TemplateBundles.First(x => x.ShowName.Equals("ProjectMappedFolderHierarchy")).HasProjectMappedFolderHierarchy);
        }

        /// <summary>
        /// Update source location
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "templateManager"), TestMethod]
        public void TemplateManagerTest_UpdateSourceLocation_ShouldUpdateCache()
        {
            var cachedFile = Path.Combine(TemplateManager.RootCacheDirectory, _guid.ToString(), "Legacy(2010).w2t");
            var sourceFile = Path.Combine(_testLocation, "Legacy(2010).w2t");
            File.SetAttributes(sourceFile, FileAttributes.ReadOnly);

            // make sure the bundle is already in the cache
            TemplateManagerTest_AddTemplateBundle_ShouldAddTemplateToListShouldCreateCacheEntry();

            // alter bundle
            File.Create(Path.Combine(_testLocation, "test.xml")).Close();
            File.SetAttributes(sourceFile, FileAttributes.Normal);
            File.WriteAllText(sourceFile, File.ReadAllText(sourceFile).Replace("Legacy(2010)", "Legacy(2010)_Foo"));
            File.SetAttributes(sourceFile, FileAttributes.ReadOnly);

            // load a new template manager
            var templateManager = new TemplateManager();
            Assert.IsNotNull(templateManager);

            // make sure new file is added and existing file is replaced even when it was readonly
            Assert.IsTrue(File.ReadAllText(cachedFile).Contains("Legacy(2010)_Foo"));
            Assert.IsTrue(File.Exists(Path.Combine(TemplateManager.RootCacheDirectory, _guid.ToString(), "test.xml")));
        }
        #endregion

    }
}
