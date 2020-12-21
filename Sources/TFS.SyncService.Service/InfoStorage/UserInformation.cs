using System;
using System.ComponentModel;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using System.Globalization;
using AIT.TFS.SyncService.Service.Properties;

namespace AIT.TFS.SyncService.Service.InfoStorage
{
    /// <summary>
    /// Class that represents information about an event that occurred during a lengthy operation like publishing.
    /// </summary>
    public class UserInformation : IUserInformation
    {
        private DateTime _occuredAt;
        private string _text;
        private string _explanation;
        private UserInformationType _type;
        private IWorkItem _source;
        private IWorkItem _destination;
        private Action _navigateToSourceAction;

        /// <summary>
        /// Raised when a property of this information has changed
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Create new instance.
        /// </summary>
        public UserInformation()
        {
            _type = UserInformationType.Warning;
            OccurredAt = DateTime.Now;
        }

        /// <summary>
        /// Create new instance for work item related information.
        /// </summary>
        /// <param name="sourceWorkItem">work item that this information is about</param>
        public UserInformation(IWorkItem sourceWorkItem)
            : this()
        {
            Source = sourceWorkItem;

            if (Source.Title.Length < 20)
            {
                Text = string.Format(CultureInfo.InvariantCulture, Resources.UserInformation_Header, Source.Title, Source.Id);
            }
            else
            {
                Text = string.Format(CultureInfo.InvariantCulture, Resources.UserInformation_HeaderShort, Source.Title.Substring(0, 20), Source.Id);
            }
        }

        /// <summary>
        /// Gets or sets the time the event occurred at.
        /// </summary>
        public DateTime OccurredAt
        {
            get { return _occuredAt; }
            set
            {
                _occuredAt = value;
                OnPropertyChanged("OccuredAt");
            }
        }

        /// <summary>
        /// Gets or sets the text.
        /// </summary>
        public string Text
        {
            get 
            { 
                return _text;
            }
            set
            {
                _text = value;
                OnPropertyChanged("Text");
            }
        }

        /// <summary>
        /// Gets or sets the explanation.
        /// </summary>
        public string Explanation
        {
            get { return _explanation; }
            set
            {
                _explanation = value;
                OnPropertyChanged("Explanation");
            }
        }

        /// <summary>
        /// Gets or sets the information type.
        /// </summary>
        public UserInformationType Type
        {
            get { return _type; }
            set
            {
                _type = value;
                OnPropertyChanged("Type");
            }
        }

        /// <summary>
        /// Gets or sets the source.
        /// </summary>
        public IWorkItem Source
        {
            get { return _source; }
            set
            {
                _source = value;
                OnPropertyChanged("Source");
            }
        }

        /// <summary>
        /// Gets or sets the destination.
        /// </summary>
        public IWorkItem Destination
        {
            get { return _destination; }
            set
            {
                _destination = value;
                OnPropertyChanged("Destination");
            }
        }

        /// <summary>
        /// Gets or sets the action to be invoked to present the user the error source
        /// </summary>
        public Action NavigateToSourceAction
        {
            get { return _navigateToSourceAction; }
            set 
            {
                _navigateToSourceAction = value;
                OnPropertyChanged("NavigateToSourceAction");
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        public override string ToString()
        {
            return "UserInformation (" +
                "Type: " + Type + "; " +
                "Occured at: " + _occuredAt.ToString("yyyy-MM-dd HH:mm:ss") + "; " +
                "Message: " + _text + "; " +
                (_source!=null ? ("Source Work Item: " + _source.Id + "; ") : string.Empty) +
                (_destination!=null ? ("Destination Work Item: " + _destination.Id + "; ") : string.Empty) +
                "Explanation:" + Explanation + 
                ")";
        }
    }
}