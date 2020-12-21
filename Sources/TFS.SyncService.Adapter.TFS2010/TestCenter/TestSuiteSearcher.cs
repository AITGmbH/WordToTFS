#region Usings
using System;
using System.Collections.Generic;
using System.Linq;
using AIT.TFS.SyncService.Adapter.TFS2012.Properties;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Factory;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    public class TestSuiteSearcher : ITestSuiteSearcher
    {
        #region Fields
        private readonly ITfsTestPlan _tfsTestPlan;
        public const char ReportingSuitePathDelimiter = '/';
        private string _originalTestPath;
        private ITfsTestSuite _suite;
        #endregion

        #region Constructor
        public TestSuiteSearcher(ITfsTestPlan testPlan)
        {
            _tfsTestPlan = testPlan;
        }
        #endregion

        #region Public methods

        /// <summary>
        /// Searches for a test suite within the test suite hiearchy below the test plan.
        /// Throws <exception cref="ArgumentException">  if, suite name or path cannot be found.</exception>
        /// </summary>
        /// <param name="suiteNameOrPath">The suite name or path which should be found.</param>
        /// <returns>The found test suite.</returns>
        public ITfsTestSuite SearchTestSuiteWithinTestPlan(string suiteNameOrPath)
        {
            return SearchTestSuite(_tfsTestPlan.RootTestSuite.TestSuites, suiteNameOrPath);
        }
        #endregion

        #region Private methods

        private ITfsTestSuite SearchTestSuite(IEnumerable<ITfsTestSuite> suites, string testSuitePath)
        {
            string potentialErrorMessage;

            if (testSuitePath.Contains('/'))
            {
                _originalTestPath = testSuitePath;
                _suite = SearchTestSuiteByPath(suites, testSuitePath);
                potentialErrorMessage = Resources.TestSuitePathCouldNotBeFound;
            }
            else
            {
                _suite = SearchtestSuiteByName(suites, testSuitePath);
                potentialErrorMessage = Resources.TestSuiteNameCouldNotBeFound;
            }

            if (_suite == null)
            {
                SyncServiceTrace.D(Resources.TestSuiteNotExist);
                throw new ArgumentException(string.Format(potentialErrorMessage, testSuitePath));
            }

            return _suite;
        }

        private ITfsTestSuite SearchtestSuiteByName(IEnumerable<ITfsTestSuite> suites, string testSuiteName)
        {
            SyncServiceTrace.D(Resources.SearchingTestSuiteByName);
            if (_tfsTestPlan.RootTestSuite.Title == testSuiteName)
            {
                return _tfsTestPlan.RootTestSuite;
            }
            var testSuites = suites as List<ITfsTestSuite>;

            if (testSuites == null)
            {
                SyncServiceTrace.D(Resources.TestSuitesNotExist);
                return null;
            }

            foreach (var suite in testSuites)
            {
                if (suite.Title == testSuiteName)
                {
                    return suite;
                }
                else if (suite.TestSuites.Any())
                {
                    var foundSuite = SearchtestSuiteByName(suite.TestSuites, testSuiteName);
                    if (foundSuite != null)
                    {
                        return foundSuite;
                    }
                }
            }

            // fallback return, if none of the conditions above was true, no approprate test suite was fond
            return null;
        }

        private ITfsTestSuite SearchTestSuiteByPath(IEnumerable<ITfsTestSuite> suites, string suitePathtestSuite)
        {
            //Example: RootSuite/SubSuite1/Sub1Sub1
            SyncServiceTrace.D(Resources.SearchingTestSuiteByPath);
            var splitted = suitePathtestSuite.Split(ReportingSuitePathDelimiter);
            if (splitted.Length > 2)
            {
                foreach (var suite in suites)
                {
                    if (suite.Title.Equals(splitted[1]))
                    {
                        var nextPart = string.Join(ReportingSuitePathDelimiter.ToString(), splitted.Where((x, index) => index >= 1));
                        var result = SearchTestSuiteByPath(suite.TestSuites, nextPart);
                        if (result != null)
                        {
                            return result;
                        }
                    }
                }
            }
            else
            {
                switch (splitted.Length)
                {
                    case 2:
                        {
                            var determinedSuite = suites.FirstOrDefault(s => s.Title == splitted[1]);
                            if (determinedSuite == null)
                            {
                                throw new ArgumentException(string.Format(Resources.TestSuitePathCouldNotBeFound, suitePathtestSuite));
                            }
                            return determinedSuite;
                        }
                    default:
                        {
                            throw new NotSupportedException(Resources.TestSuiteSearchError);
                        }
                }
            }

            return null;
        }
        #endregion
    }
}
