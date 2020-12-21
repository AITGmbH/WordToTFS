//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using AIT.TFS.SyncService.Contracts.BuildCenter;

//namespace AIT.TFS.SyncService.Model.TestReport
//{

//    /// <summary>
//    /// Helper class to print out all TestConfigurations
//    /// </summary>
//    class TestConfigurationHelper
//    {



//       public GetTestConfigurationsForTestPlanWithTestResults (ITfsServerBuild serverBuild)
//        {
            

             

//                if (serverBuild != null && serverBuild.BuildNumber == AllServerBuilds.AllServerBuildsId)
//                    // null means use all server builds
//                    serverBuild = null;

//                if (_lastUsedSelectedServerBuildForTestConfigurations != serverBuild
//                    || _lastUsedTestPlanForTestConfigurations != SelectedTestPlan)
//                {
//                    _lastUsedTestPlanForTestConfigurations = SelectedTestPlan;
//                    _testConfigurations = null;
//                    if (SelectedTestPlan == null)
//                        return _testConfigurations;

//                    _lastUsedSelectedServerBuildForTestConfigurations = serverBuild;
//                    StartBackgroundWorker(false, () =>
//                    {
//                        _testConfigurations = new List<ITfsTestConfiguration> { new AllTestConfigurations() };
//                        try
//                        {
//                            foreach (var element in TestAdapter.GetTestConfigurationsForTestPlanWithTestResults(serverBuild, SelectedTestPlan))
//                                _testConfigurations.Add(element);
//                        }
//                        finally
//                        {
//                            OnPropertyChanged("TestConfigurations");
//                            ViewDispatcher.BeginInvoke(new Action(() =>
//                            {
//                                var testConfigurationToSelect = _testConfigurations[0];
//                                if (!string.IsNullOrEmpty(StoredSelectedTestConfiguration))
//                                {
//                                    if (_testConfigurations.Any(conf => conf.Name == StoredSelectedTestConfiguration))
//                                        testConfigurationToSelect = _testConfigurations.First(conf => conf.Name == StoredSelectedTestConfiguration);
//                                }
//                                SelectedTestConfiguration = testConfigurationToSelect;
//                            }));
//                        }
//                    });
//                }
//                return _testConfigurations;

//        }
         
          
            
//    }
//}
