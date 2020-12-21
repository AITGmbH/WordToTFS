using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIT.TFS.SyncService.Contracts.Adapter
{
    public interface IViewDispatcher
    {
        bool IsDispatching { get; }
        void Invoke(Action callback);
        void BeginInvoke(Delegate method, params object[] args);
    }
}
