#region Usings
using AIT.TFS.SyncService.Model.WindowModelBase;
#endregion
namespace AIT.TFS.SyncService.Model.WindowModel
{
    /// <summary>
    /// Data driven ViewModel for generic data to be selectable in the View.
    /// </summary>
    /// <typeparam name="T">Type of the wrapped data</typeparam>
    public class DataItemModel<T> : ExtBaseModel
    {
        #region Private fields
        private bool _isChecked;
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="DataItemModel{T}"/> class.
        /// </summary>
        /// <param name="item">The item to wrap.</param>
        public DataItemModel(T item)
        {
            Item = item;
        }
        #endregion
        #region Public properties
        /// <summary>
        /// Gets the wrapped item.
        /// </summary>
        public T Item { get; private set; }

        /// <summary>
        /// Gets or sets whether item is checked.
        /// </summary>
        public bool IsChecked
        {
            get
            {
                return _isChecked;
            }
            set
            {
                _isChecked = value;
                OnPropertyChanged(nameof(IsChecked));
            }
        }
        #endregion
    }
}