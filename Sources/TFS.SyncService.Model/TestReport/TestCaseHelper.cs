#region Usings
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using AIT.TFS.SyncService.Contracts.BuildCenter;
using AIT.TFS.SyncService.Contracts.Enums.Model;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Model.Helper;
#endregion

namespace AIT.TFS.SyncService.Model.TestReport
{
    /// <summary>
    /// Help class for work with test case list.
    /// </summary>
    internal class TestCaseHelper
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="TestCaseHelper"/> class.
        /// </summary>
        /// <param name="testCases">Collection of test cases - data source.</param>
        /// <param name="pathType">Path of test cases to use by all operations in this class.
        /// (Supported is <see cref="DocumentStructureType.IterationPath"/> and <see cref="DocumentStructureType.AreaPath"/>.)</param>
        public TestCaseHelper(IEnumerable<ITfsTestCaseDetail> testCases, DocumentStructureType pathType)
        {
            if (testCases == null)
                throw new ArgumentNullException("testCases");
            CheckTestResults = false;
            TestAdapter = null;
            TestConfiguration = null;
            TestCases = testCases;
            PathType = pathType;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TestCaseHelper"/> class.
        /// </summary>
        /// <param name="testAdapter">Adapter to obtain information.</param>
        /// <param name="testPlan"><see cref="ITfsTestPlan"/> is the source of test cases to write out.</param>
        /// <param name="testConfiguration">Configuration filter used to determine test results for test case. <c>null</c> - all configurations.</param>
        /// <param name="serverBuild">Build filter used to determine test result for test case.</param>
        /// <param name="testCases">Collection of test cases - data source.</param>
        /// <param name="pathType">Path of test cases to use by all operations in this class.
        /// (Supported is <see cref="DocumentStructureType.IterationPath"/> and <see cref="DocumentStructureType.AreaPath"/>.)</param>
        public TestCaseHelper(ITfsTestAdapter testAdapter, ITfsTestPlan testPlan, ITfsTestConfiguration testConfiguration, ITfsServerBuild serverBuild,
            IEnumerable<ITfsTestCaseDetail> testCases, DocumentStructureType pathType)
        {
            if (testAdapter == null)
                throw new ArgumentNullException("testAdapter");
            if (testCases == null)
                throw new ArgumentNullException("testCases");
            CheckTestResults = true;
            TestAdapter = testAdapter;
            TestPlan = testPlan;
            TestConfiguration = testConfiguration;
            ServerBuild = serverBuild;
            TestCases = testCases;
            PathType = pathType;
        }
        #endregion
        #region Properties
        /// <summary>
        /// Gets the flag if the test results should be checked for the test cases.
        /// </summary>
        public bool CheckTestResults { get; private set; }

        /// <summary>
        /// Gets the <see cref="ITfsTestAdapter"/> to obtain the information.
        /// </summary>
        public ITfsTestAdapter TestAdapter { get; private set; }

        /// <summary>
        /// Gets the <see cref="ITfsTestPlan"/> to obtain associated information.
        /// </summary>
        public ITfsTestPlan TestPlan { get; private set; }

        /// <summary>
        /// Gets the <see cref="ITfsTestConfiguration"/> to filter the test results for test case.
        /// </summary>
        public ITfsTestConfiguration TestConfiguration { get; private set; }

        /// <summary>
        /// Gets the <see cref="ITfsServerBuild"/> to filter the test results for test case.
        /// </summary>
        public ITfsServerBuild ServerBuild { get; private set; }

        /// <summary>
        /// Gets the collection of test cases - data source.
        /// </summary>
        public IEnumerable<ITfsTestCaseDetail> TestCases { get; private set; }

        /// <summary>
        /// Gets the path of test cases to use by all operations in this class.
        /// (Supported is <see cref="DocumentStructureType.IterationPath"/> and <see cref="DocumentStructureType.AreaPath"/>.)
        /// </summary>
        public DocumentStructureType PathType { get; private set; }
        #endregion
        #region Public methods
        /// <summary>
        /// The method  creates tree structure of the test cases by <see name="PathType"/> and returns the structure.
        /// </summary>
        /// <param name="skipLevels">Count of levels to skip.</param>
        /// <returns>Required structure.</returns>
        public IEnumerable<PathElement> GetPathElements(int skipLevels)
        {
            var paths = ExtractAllPaths(PathType);
            if (paths == null)
                return null;
            return CreateTargetPathTree(CreatePathTree(paths), skipLevels);
        }

        /// <summary>
        /// The method returns all test cases as sorted list.
        /// </summary>
        /// <param name="sort">Sort to use for the returned test cases.</param>
        /// <returns>Sorted </returns>
        public IList<ITfsTestCaseDetail> GetTestCases(TestCaseSortType sort)
        {
            var testCases = new List<ITfsTestCaseDetail>();
            testCases.AddRange(TestCases);

            if (sort == TestCaseSortType.None)
                return testCases;

            testCases.Sort((a, b) =>
            {
                switch (sort)
                {
                    case TestCaseSortType.AreaPath:
                        return string.Compare(a.AreaPath, b.AreaPath, CultureInfo.CurrentCulture, CompareOptions.OrdinalIgnoreCase);
                    case TestCaseSortType.IterationPath:
                        return string.Compare(a.IterationPath, b.IterationPath, CultureInfo.CurrentCulture, CompareOptions.OrdinalIgnoreCase);
                }
                return a.WorkItemId - b.WorkItemId;
            });
            return testCases;
        }

        /// <summary>
        /// The method finds all test cases by the path in given <see cref="PathElement"/> as sorted list.
        /// </summary>
        /// <param name="pathElement"><see cref="PathElement"/> determines path of test cases.</param>
        /// <param name="sort">Sort to use for the returned test cases.</param>
        /// <returns>Sorted </returns>
        public IList<ITfsTestCaseDetail> GetTestCases(PathElement pathElement, TestCaseSortType sort)
        {
            if (pathElement == null)
                return null;
            var testCases = new List<ITfsTestCaseDetail>();
            foreach (var testCase in TestCases)
            {
                if ((PathType == DocumentStructureType.AreaPath && pathElement.WholePath == testCase.AreaPath)
                    || (PathType == DocumentStructureType.IterationPath && pathElement.WholePath == testCase.IterationPath))
                {
                    testCases.Add(testCase);
                }
            }

            if (sort == TestCaseSortType.None)
                return testCases;


            testCases.Sort((a, b) =>
            {
                switch (sort)
                {
                    case TestCaseSortType.AreaPath:
                        return string.Compare(a.AreaPath, b.AreaPath, CultureInfo.CurrentCulture, CompareOptions.OrdinalIgnoreCase);
                    case TestCaseSortType.IterationPath:
                        return string.Compare(a.IterationPath, b.IterationPath, CultureInfo.CurrentCulture, CompareOptions.OrdinalIgnoreCase);
                }
                return a.WorkItemId - b.WorkItemId;
            });
            return testCases;
        }
        #endregion
        #region Private methods
        /// <summary>
        /// The method extracts all paths (Area or Iteration) from given test cases.
        /// </summary>
        /// <param name="pathType">Path to extract - AreaPath or IterationPath.</param>
        /// <returns>All extracted paths from given <see cref="ITfsTestCaseDetail"/>.</returns>
        private IEnumerable<string> ExtractAllPaths(DocumentStructureType pathType)
        {
            if (pathType != DocumentStructureType.AreaPath && pathType != DocumentStructureType.IterationPath)
                return null;
            var paths = new List<string>();
            foreach (var testCase in TestCases)
            {
                if (pathType == DocumentStructureType.AreaPath)
                {
                    // Valid path? Already inserted?
                    if (string.IsNullOrEmpty(testCase.AreaPath) || paths.Contains(testCase.AreaPath))
                        continue;
                    // Contains test results?
                    if (CheckTestResults && !TestAdapter.TestResultExists(TestPlan, null, testCase, TestConfiguration, ServerBuild))
                        continue;
                    // Add
                    paths.Add(testCase.AreaPath);
                }
                else if (pathType == DocumentStructureType.IterationPath)
                {
                    // Valid path? Already inserted?
                    if (string.IsNullOrEmpty(testCase.IterationPath) || paths.Contains(testCase.IterationPath))
                        continue;
                    // Contains test results?
                    if (CheckTestResults && !TestAdapter.TestResultExists(TestPlan, null, testCase, TestConfiguration, ServerBuild))
                        continue;
                    // Add
                    paths.Add(testCase.IterationPath);
                }
            }
            return paths;
        }
        #endregion
        #region Private static methods
        /// <summary>
        /// The method creates tree from path strings - used for AreaPath and IterationPath
        /// </summary>
        /// <param name="paths">All paths to use to create tree.</param>
        /// <returns>Created tree from given paths.</returns>
        private static IEnumerable<PathElement> CreatePathTree(IEnumerable<string> paths)
        {
            var retVal = new List<PathElement>();
            if (paths == null)
                return retVal;
            foreach (var path in paths)
            {
                var parts = new List<string>(path.Split(PathElement.ConstDelimiter[0]));
                var firstPart = parts[0];
                parts.RemoveAt(0);
                string newParts = string.Empty;
                if (parts.Count > 0)
                    newParts = string.Join(PathElement.ConstDelimiter, parts.ToArray());
                var child = retVal.FirstOrDefault(pe => pe.PathPart == firstPart);
                if (child == null)
                {
                    var pe = new PathElement(null, firstPart);
                    pe.Add(newParts);
                    retVal.Add(pe);
                }
                else
                {
                    child.Add(newParts);
                }
            }
            retVal.Sort((a, b) =>
                string.Compare(a.PathPart, b.PathPart, CultureInfo.CurrentCulture, CompareOptions.OrdinalIgnoreCase));
            return retVal;
        }

        /// <summary>
        /// The method creates new try by given tree and skip levels number.
        /// </summary>
        /// <param name="pathElements">Source tree.</param>
        /// <param name="skipLevels">Levels to skip.</param>
        /// <returns>Create new tree.</returns>
        private static IEnumerable<PathElement> CreateTargetPathTree(IEnumerable<PathElement> pathElements , int skipLevels )
        {
            if (pathElements == null)
                return null;
            var retValue = new List<PathElement>();
            foreach (var pathElement in pathElements)
            {
                if (skipLevels < pathElement.Level)
                    retValue.Add(pathElement);
                retValue.AddRange(CreateTargetPathTree(pathElement.Childs, skipLevels));
            }
            return retValue;
        }
        #endregion
    }
}
