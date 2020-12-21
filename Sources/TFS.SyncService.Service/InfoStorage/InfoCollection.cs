using System.Collections.ObjectModel;
using AIT.TFS.SyncService.Contracts.InfoStorage;

namespace AIT.TFS.SyncService.Service.InfoStorage
{
    /// <summary>
    /// Info Collection. This collection overrides the collection changed event and invokes
    /// each subscriber with its dispatcher. This is useful if for example views are bound directly to it.
    /// </summary>
    /// <typeparam name="T">Type of the <see cref="InfoCollection{T}"/></typeparam>
    public class InfoCollection<T> : ObservableCollection<T>, IInfoCollection<T>
    {
    }
}