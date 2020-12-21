using System.Globalization;
using AIT.TFS.SyncService.Contracts.Configuration;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Linq;

namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects
{
    /// <summary>
    /// When displaying synchronization states, work items must be compared to earlier versions (revisions) of the TFS work item.
    /// This class wraps a field of a work item revision other than the current revision.
    /// It implements no setters and reads from revision values instead of the most recent work item version.
    /// </summary>
    public class TfsRevisionField : TfsField
    {
        private readonly Revision _revision;

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsRevisionField"/> class.
        /// </summary>
        /// <param name="workItem">The work item that is wrapped</param>
        /// <param name="revision">Revision to wrap.</param>
        /// <param name="configurationFieldItem">Configuration of the field.</param>
        internal TfsRevisionField(TfsWorkItem workItem, Revision revision, IConfigurationFieldItem configurationFieldItem)
            : base(workItem, configurationFieldItem)
        {
            _revision = revision;
        }

        /// <summary>
        /// The value of the current field.
        /// </summary>
        public override string Value
        {
            get
            {
                return Convert.ToString(_revision.Fields[ReferenceName].Value, CultureInfo.CurrentCulture);
            }
            set
            {
            }
        }

        /// <summary>
        /// Get attachment by name.
        /// </summary>
        /// <param name="name">Name of the attachment</param>
        /// <returns>The attachment if it exists or null.</returns>
        protected override Attachment GetAttachment(string name)
        {
            return
                _revision.Attachments.Cast<Attachment>().FirstOrDefault(
                    attachment => name.Equals(attachment.Name, StringComparison.OrdinalIgnoreCase));
        }
    }
}