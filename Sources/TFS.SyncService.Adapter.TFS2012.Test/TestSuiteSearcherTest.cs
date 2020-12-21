using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace TFS.SyncService.Adapter.TFS2012.Test
{
    using System.Collections.Generic;
    using System.Diagnostics;
    using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
    using AIT.TFS.SyncService.Contracts.TestCenter;

    [TestClass]
    public class TestSuiteSearcherTest
    {

        /// <summary>
        /// Creates a TestPlan with a small tree structure
        /// RootSuite/SubSuite_1/SubSuite_1_1
        /// RootSuite/SubSuite_1/SubSuite_1_2
        /// RootSuite/SubSuite_2/SubSuite_2_1
        /// RootSuite/SubSuite_2/SubSuite_2_2
        /// </summary>
        /// <returns></returns>
        private static ITfsTestPlan GetPreparedTestPlanWithSuites()
        {
            var subList = new List<ITfsTestSuite>();
            var rootSuite = new Mock<ITfsTestSuite>();
            rootSuite.SetupGet(x => x.Title).Returns("RootSuite");
            rootSuite.SetupGet(x => x.TestSuites).Returns(subList);

            for (var i = 1; i < 3; i++)
            {
                var subSuiteLvl1 = new Mock<ITfsTestSuite>();
                subSuiteLvl1.SetupGet(x => x.Title).Returns("SubSuite_" + i);
                var subListLvl2 = new List<ITfsTestSuite>();
                subSuiteLvl1.SetupGet(x => x.TestSuites).Returns(subListLvl2);

                for (var j = 1; j < 3; j++)
                {
                    var subSuiteLvl2 = new Mock<ITfsTestSuite>();
                    subSuiteLvl2.SetupGet(x => x.Title).Returns("SubSuite_" + i + "_" + j);
                    subListLvl2.Add(subSuiteLvl2.Object);
                }

                subList.Add(subSuiteLvl1.Object);
            }
            //Test(rootSuite.Object);

            var myPlan = new Mock<ITfsTestPlan>();
            myPlan.SetupGet(x => x.RootTestSuite).Returns(rootSuite.Object);

            return myPlan.Object;
        }

        private void Test(ITfsTestSuite rootSuite)
        {
            Debug.WriteLine(rootSuite.Title);
            foreach (var suite in rootSuite.TestSuites)
            {
                Test(suite);
            }
        }

        [TestMethod]
        public void TestSuiteSearcher_SearchTestSuiteWithinTestPlan_ByName()
        {
            var searchTitle = "SubSuite_1";
            var myPreparedTestPlan = GetPreparedTestPlanWithSuites();
            var mySearcher = new TestSuiteSearcher(myPreparedTestPlan);
            var foundSuite = mySearcher.SearchTestSuiteWithinTestPlan(searchTitle);

            Assert.IsTrue(foundSuite.Title == searchTitle);
        }

        [TestMethod]
        public void TestSuiteSearcher_SearchTestSuiteWithinTestPlan_ByNameDeep()
        {
            var searchTitle = "SubSuite_2_2";
            var myPreparedTestPlan = GetPreparedTestPlanWithSuites();
            var mySearcher = new TestSuiteSearcher(myPreparedTestPlan);
            var foundSuite = mySearcher.SearchTestSuiteWithinTestPlan(searchTitle);

            Assert.IsTrue(foundSuite.Title == searchTitle);
        }

        [TestMethod]
        public void TestSuiteSearcher_SearchTestSuiteWithinTestPlan_ByPath()
        {
            var searchPath = "RootSuite/SubSuite_2/SubSuite_2_2";
            var myPreparedTestPlan = GetPreparedTestPlanWithSuites();
            var mySearcher = new TestSuiteSearcher(myPreparedTestPlan);
            var foundSuite = mySearcher.SearchTestSuiteWithinTestPlan(searchPath);

            Assert.IsTrue(foundSuite.Title == "SubSuite_2_2");
        }
    }
}
