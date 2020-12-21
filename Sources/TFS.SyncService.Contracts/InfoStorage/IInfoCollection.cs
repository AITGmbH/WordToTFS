using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;

namespace AIT.TFS.SyncService.Contracts.InfoStorage
{
    /// <summary>
    /// Interface defines the functionality of a list - list of information that generated during publishing process.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IInfoCollection<T> : IList<T>, ICollection<T>, IEnumerable<T>, INotifyCollectionChanged,
                                          INotifyPropertyChanged
    {
    }
}