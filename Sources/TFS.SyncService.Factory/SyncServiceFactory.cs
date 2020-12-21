using System;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Adapter;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Factory
{
    /// <summary>
    /// Factory class provides adapter and service interfaces.
    /// </summary>
    public static class SyncServiceFactory
    {
        private static readonly Dictionary<Type, object> Services = new Dictionary<Type, object>();

        /// <summary>
        /// Method gets the interface of new object that implements
        /// the <see cref="IWorkItemSyncAdapter"/> interface for word document - table implementation.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">Configuration to use</param>
        /// <returns>Created object.</returns>
        public static IWordSyncAdapter CreateWord2007TableWorkItemSyncAdapter(object document, IConfiguration configuration)
        {
            var service = GetService<IWord2007AdapterCreator>();
            if (service == null)
                return null;
            return service.CreateTableAdapter(document, configuration);
        }

        /// <summary>
        /// Method gets the interface of new object that implements
        /// the <see cref="IWord2007TestReportAdapter"/> interface for word document - test report implementation.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">Configuration to use.</param>
        /// <returns>Created object.</returns>
        public static IWord2007TestReportAdapter CreateWord2007TestReportAdapter(object document, IConfiguration configuration)
        {
            var service = GetService<IWord2007AdapterCreator>();
            return service == null ? null : service.CreateTestReportAdapter(document, configuration);
        }

        /// <summary>
        /// Method gets the interface of new object that implements
        /// the <see cref="IWord2007TestReportAdapter"/> interface for word document - test report implementation.
        /// </summary>
        /// <param name="document">Word document to work with.</param>
        /// <param name="configuration">Configuration to use.</param>
        /// <param name="testSuite">The test suite.</param>
        /// <param name="allTestCases">All test cases that test suite contains.</param>
        /// <returns>Created object.</returns>
        public static IWord2007TestReportAdapter CreateWord2007TestReportAdapter(object document, IConfiguration configuration, ITfsTestSuite testSuite, IList<ITfsTestCaseDetail> allTestCases)
        {
            var service = GetService<IWord2007AdapterCreator>();
            return service == null ? null : service.CreateTestReportAdapter(document, configuration, testSuite, allTestCases);
        }

        /// <summary>
        /// Method gets the interface of new object that implements the <see cref="IWorkItemSyncAdapter"/> interface for TFS 2008.
        /// </summary>
        /// <param name="serverName">Name of the TFS server to work with.</param>
        /// <param name="projectName">Name of the TFS project to work with.</param>
        /// <param name="queryConfiguration">Configuration describing what work items to load.</param>
        /// <param name="configuration">Configuration that applies while opening the adapter.</param>
        /// <returns>Created object.</returns>
        public static IWorkItemSyncAdapter CreateTfs2008WorkItemSyncAdapter(string serverName, string projectName, IQueryConfiguration queryConfiguration, IConfiguration configuration)
        {
            var service = GetService<ITfsAdapterCreator>();
            if (service == null)
                return null;
            return service.CreateAdapter(serverName, projectName, queryConfiguration, configuration);
        }

        /// <summary>
        /// Methods gets the interface of new object that implements the <see cref="IWorkItemSyncAdapter"/> interface.
        /// </summary>
        /// <param name="serverName">Name of the TFS server to work with.</param>
        /// <param name="projectName">Name of the TFS project to work with.</param>
        /// <param name="configuration">Configuration that applies while opening the adapter.</param>
        /// <returns></returns>
        public static ITfsService CreateTfsService(string serverName, string projectName, IConfiguration configuration)
        {
            return CreateTfs2008WorkItemSyncAdapter(serverName, projectName, null, configuration) as ITfsService;
        }

        /// <summary>
        /// Methods gets the interface of new object that implements the <see cref="ITfsTestAdapter"/> interface.
        /// </summary>
        /// <param name="serverName">Name of the TFS server to work with.</param>
        /// <param name="projectName">Name of the TFS project to work with.</param>
        /// <param name="configuration">Configuration to use.</param>
        /// <returns>Method returns created object that implements <see cref="ITfsTestAdapter"/></returns>
        public static ITfsTestAdapter CreateTfsTestAdapter(string serverName, string projectName, IConfiguration configuration)
        {
            return CreateTfs2008WorkItemSyncAdapter(serverName, projectName, null, configuration) as ITfsTestAdapter;
        }

        /// <summary>
        /// Register a new service.
        /// </summary>
        /// <typeparam name="T">Type of service.</typeparam>
        /// <param name="instance">Instance of the service.</param>
        public static void RegisterService<T>(T instance)
        {
            if (Services.ContainsKey(typeof(T)))
            {
                Services[typeof(T)] = instance;
            }
            else
            {
                Services.Add(typeof(T), instance);
            }
        }

        /// <summary>
        /// Gets the registered service.
        /// </summary>
        /// <typeparam name="T">Type of the service to get.</typeparam>
        /// <returns>Registered instance of the service of the desired type.</returns>
        public static T GetService<T>()
        {
            if (Services.ContainsKey(typeof(T)))
                return (T)Services[typeof(T)];
            return default(T);
        }
    }
}