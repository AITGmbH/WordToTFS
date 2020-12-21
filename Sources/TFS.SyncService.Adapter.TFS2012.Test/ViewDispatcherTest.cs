using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFS.SyncService.Adapter.TFS2012.Test
{
    using System.Windows.Threading;
    using AIT.TFS.SyncService.Adapter.Word2007;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class ViewDispatcherTest
    {
        [TestMethod]
        public void ViewDispatcherInitializedWithNull()
        {
            var invokeWasCalled = false;
            var beginInvokeWasCalled = false;
            var viewDispatcher = new ViewDispatcher(null);

            viewDispatcher.Invoke(()=> invokeWasCalled = true);
            viewDispatcher.BeginInvoke(new Action(() => beginInvokeWasCalled = true));

            Assert.IsFalse(viewDispatcher.IsDispatching);
            Assert.IsTrue(invokeWasCalled);
            Assert.IsTrue(beginInvokeWasCalled);
        }
    }
}
