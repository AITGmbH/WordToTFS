using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIT.TFS.SyncService.Adapter.Word2007
{
    using System.Windows.Threading;
    using AIT.TFS.SyncService.Contracts.Adapter;
    public class ViewDispatcher: IViewDispatcher
    {
        private readonly Dispatcher _dispatcher;

        public bool IsDispatching
        {
            get
            {
                return _dispatcher != null;
            }
        }

        public ViewDispatcher(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
        }

        public void Invoke(Action callback)
        {
            if (_dispatcher != null)
            {
                _dispatcher.Invoke(callback);
            }
            else
            {
                callback.Invoke();
            }
        }

        public void BeginInvoke(Delegate method, params object[] args)
        {
            if (_dispatcher != null)
            {
                _dispatcher.BeginInvoke(method, args);
            }
            else
            {
                method.DynamicInvoke(args);
            }
        }
    }
}
