using System;
using System.ComponentModel;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;

namespace AIT.TFS.SyncService.Contracts.InfoStorage
{
    /// <summary>
    /// Interface for generic user information.
    /// </summary>
    public interface IUserInformation : INotifyPropertyChanged
    {
        /// <summary>
        /// Gets or sets the time that the error occurred.
        /// </summary>
        DateTime OccurredAt { get; set; }

        /// <summary>
        /// Gets or sets the error text.
        /// </summary>
        string Text { get; set; }

        /// <summary>
        /// Gets or sets the explanation for the error if exists.
        /// </summary>
        string Explanation { get; set; }

        /// <summary>
        /// Gets or sets the information type.
        /// </summary>
        UserInformationType Type { get; set; }

        /// <summary>
        /// Gets or sets the source.
        /// </summary>
        IWorkItem Source { get; set; }

        /// <summary>
        /// Gets or sets the destination.
        /// </summary>
        IWorkItem Destination { get; set; }

        /// <summary>
        /// Gets or sets the action to be invoked to present the user the error source
        /// </summary>
        Action NavigateToSourceAction { get; set; } 
    }
}
