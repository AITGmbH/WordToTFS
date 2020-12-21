using AIT.TFS.SyncService.Adapter.TFS2012.Properties;
using AIT.TFS.SyncService.Adapter.TFS2012.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.TfsHelper;
using AIT.TFS.SyncService.Contracts.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Service.Configuration;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects
{
    /// <summary>
    /// Class implements <see cref="IWorkItem"/> and used in TFS Adapter.
    /// </summary>
    public class TfsWorkItem : IWorkItem
    {
        #region Fields
        private readonly IEnumerable<string> _lookupFields;
        private readonly IConfigurationItem _configuration;
        private IFieldCollection _fields;
        private ITestBase _testCase;
        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsWorkItem"/> class.
        /// </summary>
        /// <param name="workItem">The TFS work item to be wrapped</param>
        /// <param name="tfsAdapter">TFS adapter.</param>
        /// <param name="lookupFields">When this is set, the work item is only used to look up field values for link expansion.</param>
        /// <param name="configuration">Configuration of the  </param>
        internal TfsWorkItem(WorkItem workItem, Tfs2012SyncAdapter tfsAdapter, IEnumerable<string> lookupFields, IConfigurationItem configuration)
        {
            if (workItem == null) throw new ArgumentNullException("workItem");
            if (tfsAdapter == null) throw new ArgumentNullException("tfsAdapter");
            if (configuration == null) throw new ArgumentNullException("configuration");

            WorkItem = workItem;
            TfsAdapter = tfsAdapter;
            _lookupFields = lookupFields;
            _configuration = configuration;
        }

        #endregion

        #region Private and internal properties

        /// <summary>
        /// Gets the associated work item.
        /// </summary>
        internal WorkItem WorkItem { get; private set; }

        /// <summary>
        /// Gets the <see cref="WorkItemStore"/> of associated work item.
        /// </summary>
        internal WorkItemStore WorkItemStore
        {
            get
            {
                if (null != WorkItem)
                    return WorkItem.Store;

                return null;
            }
        }

        /// <summary>
        /// Gets TFS adapter instance.
        /// </summary>
        internal Tfs2012SyncAdapter TfsAdapter { get; private set; }

        /// <summary>
        /// Gets Test Case if current WorkItem is Test Case type.
        /// </summary>
        internal ITestBase TestCase
        {
            get
            {
                if (_testCase == null &&
                    TfsAdapter.TestManagement != null &&
                    WorkItem != null &&
                    TfsAdapter.TestManagement.TestCases.IsWorkItemCompatible(WorkItem))
                {
                    if (Id > 0)
                    {
                        _testCase = TfsAdapter.TestManagement.TestCases.Find(Id);
                    }
                    else
                    {
                        _testCase = TfsAdapter.TestManagement.CreateFromWorkItem(WorkItem);
                    }
                }

                return _testCase;
            }
        }
        #endregion

        #region IWorkItem Members

        /// <summary>
        /// The work item id of the current work item. This value has to be unique among all work items.
        /// </summary>
        public int Id
        {
            get { return WorkItem.Id; }
        }

        /// <summary>
        /// Gets the work item revision.
        /// </summary>
        public int Revision
        {
            get
            {
                return WorkItem.Revision;
            }
        }

        /// <summary>
        /// Gets work item title.
        /// </summary>
        public string Title
        {
            get { return WorkItem.Title; }
        }

        /// <summary>
        /// Determines the type of the work item.
        /// </summary>
        public string WorkItemType
        {
            get { return WorkItem.Type.Name; }
        }

        /// <summary>
        /// All fields which are supported by the current work item.
        /// </summary>
        public IFieldCollection Fields
        {
            get
            {
                if (null == _fields)
                {
                    IList<IField> fields = new List<IField>();

                    if (_lookupFields == null)
                    {
                        var allCustomFields = new List<string>();

                        foreach (var fieldItem in _configuration.FieldConfigurations)
                        {
                            if (!WorkItem.Fields.Contains(fieldItem.ReferenceFieldName)) continue;
                            if (fieldItem.ReferenceFieldName.Equals(FieldReferenceNames.TestSteps))
                            {
                                fields.Add(new TfsTestCaseField(this, fieldItem));
                            }
                            else
                            {
                                fields.Add(new TfsField(this, fieldItem));
                                if(_configuration.FieldConfigurations.Where(f => f.ReferenceFieldName == fieldItem.OLEMarkerField).Count() > 0) continue;
                                if (fieldItem.OLEMarkerField != null && !allCustomFields.Contains(fieldItem.OLEMarkerField))
                                {
                                    allCustomFields.Add(fieldItem.OLEMarkerField);
                                }
                            }
                        }
                        
                        foreach (var customField in allCustomFields)
                        {
                            var tempConfiguration = new ConfigurationFieldItem(customField, customField, FieldValueType.PlainText, Direction.TfsToOther, 0, 0, string.Empty, false, HandleAsDocumentType.OleOnDemand, null, string.Empty, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null);
                            var tempField = new TfsField(this, tempConfiguration);

                            if (!fields.Contains(tempField))
                            {
                                fields.Add(tempField);
                            }
                        }
                    }
                    else
                    {
                        //TODO: MIS 2016-03-30: Wyh using if and custom configurationFieldItem instead of upper version (Maybe because this is for link feature wich not must have a configuration)?
                        foreach (var field in _lookupFields)
                        {
                            var currentFieldConfig = _configuration.FieldConfigurations.FirstOrDefault(x => x.ReferenceFieldName.Equals(field));
                            //fields.Add(new TfsField(this, currentFieldConfig));

                            if (currentFieldConfig != null)
                            {
                                //fields.Add(new TfsField(this, currentFieldConfig));
                                fields.Add(new TfsField(this, new ConfigurationFieldItem(field, field, currentFieldConfig.FieldValueType, Direction.TfsToOther, 0, 0, string.Empty, false, HandleAsDocumentType.OleOnDemand, null, string.Empty, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null)));
                            }
                            else
                            {
                                fields.Add(new TfsField(this, new ConfigurationFieldItem(field, field, FieldValueType.PlainText, Direction.TfsToOther, 0, 0, string.Empty, false, HandleAsDocumentType.OleOnDemand, null, string.Empty, null, ShapeOnlyWorkaroundMode.AddSpace, false, null, null, null)));
                            }

                        }
                    }

                    _fields = new TfsFieldCollection(fields);
                }

                return _fields;
            }
        }

        /// <summary>
        /// Returns whether this work item has been modified
        /// </summary>
        public bool IsDirty
        {
            get
            {
                return WorkItem.IsDirty;
            }
        }


        /// <summary>
        /// Returns whether this work item is new
        /// </summary>
        public bool IsNew
        {
            get
            {
                return WorkItem.IsNew;
            }
        }

        /// <summary>
        /// Ids of linked work items, grouped by link type
        /// </summary>
        public Dictionary<IConfigurationLinkItem, int[]> Links
        {
            get
            {
                var links = new Dictionary<IConfigurationLinkItem, int[]>();

                foreach (IConfigurationLinkItem linkItem in _configuration.Links)
                {
                    WorkItemLinkTypeEnd linkTypeEnd;
                    if (WorkItemStore.WorkItemLinkTypes.LinkTypeEnds.TryGetByName(linkItem.LinkValueType,
                                                                                  out linkTypeEnd))
                    {
                        var ids = new List<int>();

                        // Collected ids of linked work items
                        foreach (WorkItemLink workItemLink in WorkItem.WorkItemLinks)
                        {
                            if (workItemLink.LinkTypeEnd.Id == linkTypeEnd.Id)
                            {
                                ids.Add(workItemLink.TargetId);
                            }
                        }
                        links.Add(linkItem, ids.ToArray<int>());
                    }
                }

                return links;
            }
        }

        /// <summary>
        /// Not supported.
        /// </summary>
        public void CreateBookmark()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Return wheter this work items had errors during the validation on the word side
        /// </summary>
        public bool HasWordValidationError
        {
            get;
            set;
        }

        /// <summary>
        /// Adds multiple links to other work items to this work item.
        /// </summary>
        /// <param name="adapter">Adapter used to look up more information about the referenced work items.</param>
        /// <param name="ids">The ids of the other work items.</param>
        /// <param name="linkType">The reference name of the link type.</param>
        /// <param name="overwrite">Sets whether the passed links replace all existing links of the given type.</param>
        /// <returns>True if the link was successfully added, false otherwise.</returns>
        public bool AddLinks(IWorkItemSyncAdapter adapter, int[] ids, string linkType, bool overwrite)
        {
            return this.AddLinks(adapter, ids, linkType, string.Empty, overwrite);
        }

        /// <summary>
        /// Adds multiple links to other work items to this work item.
        /// </summary>
        /// <param name="adapter">Adapter used to look up more information about the referenced work items.</param>
        /// <param name="ids">The ids of the other work items.</param>
        /// <param name="linkType">The reference name of the link type.</param>
        /// <param name="linkedWorkItemTypes">Determines a number of Work Item Types that can be used to filter the links in addition to the linkValueType.</param>
        /// <param name="overwrite">Sets whether the passed links replace all existing links of the given type.</param>
        /// <returns>True if the link was successfully added, false otherwise.</returns>
        public bool AddLinks(IWorkItemSyncAdapter adapter, int[] ids, string linkType, string linkedWorkItemTypes, bool overwrite)
        {
            WorkItemLinkTypeEnd linkTypeEnd;
            if (!WorkItemStore.WorkItemLinkTypes.LinkTypeEnds.TryGetByName(linkType, out linkTypeEnd))
                return false;

            // only add links to work items that are defined in the LinkedWorkItemTypes-attribute of the link
            if (!string.IsNullOrEmpty(linkedWorkItemTypes))
            {
                // split by comma
                // after that trim each element of the string array that results from the plit operation before
                //if(linkedWorkItemTypes.Contains)
                var linkedWorkItemTypesList = linkedWorkItemTypes.Split(',').Select(s => s.Trim()).ToList();

                foreach (var id in ids)
                {
                    var type = WorkItemStore.GetWorkItem(id).Type.Name;
                    if (!linkedWorkItemTypesList.Contains(type))
                        throw new ArgumentException(string.Format(CultureInfo.CurrentCulture, Resources.LinkWarning_WorkItemTypeViolation, id, type));
                }
                ids = ids.Where(id => linkedWorkItemTypesList.Contains(WorkItemStore.GetWorkItem(id).Type.Name)).ToArray();
            }

            if (ids == null) return false;

            // Remove all links of specified type
            if (overwrite)
            {
                var lcdelete =
                    WorkItem.Links.OfType<RelatedLink>().Where(link => link.LinkTypeEnd.Id == linkTypeEnd.Id).ToList();

                foreach (var rl in lcdelete)
                {
                    WorkItem.Links.Remove(rl);
                }
            }

            return this.AddLinks(ids, linkTypeEnd);
        }

        /// <summary>
        /// Adds a link to the work items with the given ids.
        /// </summary>
        /// <param name="ids">The ids.</param>
        /// <param name="linkTypeEnd">The link type.</param>
        /// <returns></returns>
        private bool AddLinks(int[] ids, WorkItemLinkTypeEnd linkTypeEnd)
        {
            // Check for OneToMany violations (e.g. more than one parent)
            if (ids.Length > 1 && linkTypeEnd.IsForwardLink == false && linkTypeEnd.LinkType.IsOneToMany)
            {
                throw new ArgumentException(string.Format(CultureInfo.CurrentCulture, Resources.LinkError_OneToManyViolation, linkTypeEnd.Name));
            }

            // remove all links of this link type and only for ids which are going to add
            var tfsLinks = (from Link l in WorkItem.Links select l).ToArray();
            foreach (Link link in tfsLinks)
            {
                var workItemLink = link as RelatedLink;
                if (workItemLink != null)
                {
                    if (workItemLink.LinkTypeEnd.Id == linkTypeEnd.Id && ids.Contains(workItemLink.RelatedWorkItemId))
                        WorkItem.Links.Remove(link);
                }
            }

            // add new links for this type
            foreach (var id in ids)
            {
                var relatedLink = new RelatedLink(linkTypeEnd, id);
                WorkItem.Links.Add(relatedLink);
            }

            return true;
        }

        /// <summary>
        /// Refresh work item to get latest state.
        /// </summary>
        public void Refresh()
        {
            WorkItem.SyncToLatest();
        }

        /// <summary>
        /// Get revision number for this work item. If the referenceFieldName is defined then
        /// gets revision number which belongs to the field -> last revision number for certain field.
        /// </summary>
        /// <param name="referenceFieldName">Reference field name.</param>
        /// <returns></returns>
        public int GetFieldRevision(string referenceFieldName)
        {
            Guard.ThrowOnArgumentNull(referenceFieldName, "referenceFieldName");

            int fieldRevisionIndex = GetFieldValueRevision(referenceFieldName);
            int fieldAttachmentRevisionIndex = GetAttachmentRevision(StringHelper.GetMicroDocumentName(referenceFieldName));

            return Math.Max(fieldRevisionIndex, fieldAttachmentRevisionIndex);
        }

        /// <summary>
        /// Gets a list of fields and their values for a given revision of the work item.
        /// </summary>
        /// <param name="revision">Revision for which to retrieve values.</param>
        /// <returns>
        /// List of fields for this revision.
        /// </returns>
        public IFieldCollection GetWorkItemByRevision(int revision)
        {
            if (revision == Revision)
            {
                return Fields;
            }

            var fields = Fields.Select(field => new TfsRevisionField(this, WorkItem.Revisions[revision - 1], field.Configuration));
            return new TfsFieldCollection(fields);
        }

        /// <summary>
        /// Gets the configuration of this work item and all its fields
        /// </summary>
        public IConfigurationItem Configuration { get { return _configuration; } }

        /// <summary>
        /// Returns whether this and another work item are wrapper for the same work item.
        /// </summary>
        /// <param name="other">Other work item.</param>
        /// <returns>
        /// True, if both work item wrappers are wrapper for the same work item.
        /// </returns>
        public bool IsSameWorkItem(IWorkItem other)
        {
            return other == this;
        }

        /// <summary>
        /// Delete Id and Revision
        /// </summary>
        public void DeleteIdAndRevision()
        {
            throw new NotSupportedException();
        }

        #endregion

        #region Private and internal methods

        /// <summary>
        /// Method prepares the work item for write operation.
        /// </summary>
        internal void OpenWorkItemForWrite()
        {
            if (!WorkItem.IsOpen)
            {
                //WorkItem = WorkItem.Store.GetWorkItem(WorkItem.Id);
                WorkItem.Open();
            }
        }

        /// <summary>
        /// Updates all image source in all HTML fields to point to their uploaded attachment links
        /// </summary>
        internal void SetHtmlValuesWithUpdatedImageSources()
        {
            if (null != _fields)
            {
                foreach (IField field in Fields)
                {
                    var tfsField = field as TfsField;
                    if (tfsField != null)
                        tfsField.SetHTMLValueWithUpdatedImageSources();
                }
            }
        }

        /// <summary>
        /// Delete temporary files after save process.
        /// </summary>
        internal void DeleteTempFilesAfterSave()
        {
            if (_fields != null)
            {
                foreach (IField field in _fields)
                {
                    var tfsField = field as TfsField;
                    if (tfsField != null)
                        tfsField.DeleteTempFileAfterSave();
                }
            }
        }

        /// <summary>
        /// Get revision index of latest change to the given field.
        /// </summary>
        /// <param name="referenceFieldName">Field reference name.</param>
        /// <returns>Revision index.</returns>
        private int GetFieldValueRevision(string referenceFieldName)
        {
            var revisions = WorkItem.Revisions.Cast<Revision>().Where(x => x.Fields.Contains(referenceFieldName)).OrderBy(x => x.Index).Reverse().ToList();
            Revision previousRevision = revisions.ElementAt(0);
            var mostRecentValue = previousRevision.Fields[referenceFieldName].Value;

            // find the most recent revision that is different to the current value
            foreach (var revision in revisions)
            {
                var field = revision.Fields[referenceFieldName];

                if ((mostRecentValue == null && field.Value != null) ||
                    (mostRecentValue != null && mostRecentValue.Equals(field.Value) == false))
                {
                    // our indices begin with 1 not with 0
                    return previousRevision.Index + 1;
                }

                previousRevision = revision;
            }

            // value has never changed since the first revision, there is no entry that differs
            return 0;
        }

        /// <summary>
        /// Get revision number for certain attachment.
        /// </summary>
        /// <param name="attachmentName">Attachment name.</param>
        /// <returns>Revision number.</returns>
        private int GetAttachmentRevision(string attachmentName)
        {
            Revision lastChangedAttachemntRevision = null;
            Attachment previousAttachment = null;
            // must search for attachment changes between revisions.
            foreach (Revision revision in WorkItem.Revisions)
            {
                var attachment =
                    revision.Attachments.Cast<Attachment>().FirstOrDefault(a => a.Name.Equals(attachmentName, StringComparison.OrdinalIgnoreCase));
                if (attachment != null)
                {
                    if (previousAttachment == null ||
                        attachment.Length != previousAttachment.Length ||
                        attachment.AttachedTime.Equals(previousAttachment.AttachedTime) == false)
                    {
                        lastChangedAttachemntRevision = revision;
                    }

                    previousAttachment = attachment;
                }
            }

            if (lastChangedAttachemntRevision != null)
            {
                return lastChangedAttachemntRevision.Index + 1;
            }

            return 0;
        }

        #endregion
    }
}
