#region Usings
using System.ComponentModel;
#endregion
namespace AIT.TFS.SyncService.Model.WindowModelBase
{
    /// <summary>
    /// Class defines base functionality for all model classes.
    /// All window model classes should be derived from <see cref="ExtBaseModel"/> or from <see cref="BaseModel"/>.
    /// </summary>
    public abstract class ExtBaseModel : BaseModel, INotifyPropertyChanged
    {
        #region INotifyPropertyChanged Members

        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        #region Protected properties

        /// <summary>
        /// Method triggers the <see cref="PropertyChanged"/> event.
        /// </summary>
        /// <param name="sender">Sender used in <see cref="PropertyChanged"/> event.</param>
        /// <param name="propertyName">Name of the property used in <see cref="PropertyChangedEventArgs"/>.</param>
        protected void TriggerPropertyChanged(object sender, string propertyName)
        {
            PropertyChanged?.Invoke(sender, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Method triggers the <see cref="PropertyChanged"/> event.
        /// </summary>
        /// <param name="propertyName">Name of the property used in <see cref="PropertyChangedEventArgs"/>.</param>
        protected void TriggerPropertyChanged(string propertyName)
        {
            TriggerPropertyChanged(this, propertyName);
        }

        /// <summary>
        /// Method triggers the <see cref="PropertyChanged"/> event.
        /// </summary>
        /// <param name="propertyName">Name of the property used in <see cref="PropertyChangedEventArgs"/>.</param>
        protected void OnPropertyChanged(string propertyName)
        {
            TriggerPropertyChanged(this, propertyName);
        }
 
        #endregion Protected properties
    }
}