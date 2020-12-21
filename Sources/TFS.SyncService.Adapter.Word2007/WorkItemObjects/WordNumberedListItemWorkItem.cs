using AIT.TFS.SyncService.Adapter.Word2007.Properties;
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.TfsHelper;
using AIT.TFS.SyncService.Contracts.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using Microsoft.Office.Interop.Word;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Threading;

namespace AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects
{
    /// <summary>
    /// WorkItem represents on numbered list item in the cell of one table work item.
    /// </summary>
    /// <remarks>In this work item is only title and description supported. Nothing more.</remarks>
    public class WordNumberedListItemWorkItem : IWorkItem
    {
        #region Private fields
        private readonly Range _wholeTitle;
        private readonly IConfiguration _configuration;
        private IFieldCollection _fields;
        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Constructor creates an instance of the class <see cref="WordNumberedListItemWorkItem"/>.
        /// </summary>
        /// <param name="configuration">Type of work item</param>
        /// <param name="title">Title of the work item.</param>
        /// <param name="description">Description of the work item.</param>
        /// <param name="mappingConfiguration">Configuration of the whole mapping</param>
        public WordNumberedListItemWorkItem(IConfigurationFieldToLinkedItem configuration, Range title, Range description, IConfiguration mappingConfiguration)
        {
            if (configuration == null)
                throw new ArgumentNullException("configuration");

            if(mappingConfiguration == null) throw new ArgumentNullException("mappingConfiguration");

            _wholeTitle = title;

            ConfigurationToLinkedItem = configuration;
            WorkItemType = configuration.LinkedWorkItemType;
            WorkItemDescription = description;
            _configuration = mappingConfiguration;

            PrepareTitle();
            InitializeFields();
        }
        #endregion Constructors

        #region Public properties
        /// <summary>
        /// Gets the title of the work item.
        /// </summary>
        public Range WorkItemTitle { get; private set; }

        /// <summary>
        /// Gets the first addition of title
        /// </summary>
        public Range NumberedListItemFirstAddition { get; private set; }

        /// <summary>
        /// Gets the second addition of title
        /// </summary>
        public Range NumberedListItemSecondAddition { get; private set; }

        /// <summary>
        /// Gets the third addition of title
        /// </summary>
        public Range NumberedListItemThirdAddition { get; private set; }

        /// <summary>
        /// Gets the description of the work item.
        /// </summary>
        public Range WorkItemDescription { get; private set; }
        #endregion Public properties

        #region Private properties
        private IConfigurationFieldToLinkedItem ConfigurationToLinkedItem { get; set; }
        #endregion Private properties

        #region Private methods
        private void PrepareTitle()
        {
            var lefts = new List<int>();
            var rights = new List<int>();

            var originalStart = _wholeTitle.Start;
            var originalEnd = _wholeTitle.End;

            var searching = _wholeTitle.Document.Range(_wholeTitle.Start, _wholeTitle.End);

            searching.Find.Text = "]";
            searching.Find.Forward = false;
            searching.Find.Execute();
            while (searching.Find.Found)
            {
                if (searching.End < originalStart)
                {
                    break;
                }
                rights.Add(searching.End);
                searching.Find.Execute();
            }

            searching = _wholeTitle.Document.Range(_wholeTitle.Start, _wholeTitle.End);
            searching.Find.Text = "[";
            searching.Find.Forward = false;
            searching.Find.Execute();
            while (searching.Find.Found)
            {
                if (searching.End < originalStart)
                    break;
                lefts.Add(searching.End);
                searching.Find.Execute();
            }

            if (rights.Count < 3 || lefts.Count < 3
                || rights[0] < lefts[0]
                || rights[1] < lefts[1]
                || rights[2] < lefts[2])
            {
                //var plainText = wholeTitle.Text.Replace('\r', '\n').Replace('\v', '\n').Replace('\a', '\n');
                //var trimedPlainText = plainText.TrimEnd();
                object objMoveChar = WdUnits.wdCharacter;
                object objNeg1Count = -1;// trimedPlainText.Length - plainText.Length;
                _wholeTitle.MoveEnd(ref objMoveChar, ref objNeg1Count);
                _wholeTitle.InsertAfter(Resources.NumberedListItemsTemplate);
                NumberedListItemFirstAddition = _wholeTitle.Document.Range(_wholeTitle.End - 7, _wholeTitle.End - 7);
                NumberedListItemSecondAddition = _wholeTitle.Document.Range(_wholeTitle.End - 4, _wholeTitle.End - 4);
                NumberedListItemThirdAddition = _wholeTitle.Document.Range(_wholeTitle.End - 1, _wholeTitle.End - 1);
                WorkItemTitle = _wholeTitle.Document.Range(originalStart, originalEnd - 1);//(plainText.Length - trimedPlainText.Length));
                return;
            }
            NumberedListItemThirdAddition = _wholeTitle.Document.Range(lefts[0], rights[0] - 1);
            NumberedListItemSecondAddition = _wholeTitle.Document.Range(lefts[1], rights[1] - 1);
            NumberedListItemFirstAddition = _wholeTitle.Document.Range(lefts[2], rights[2] - 1);
            WorkItemTitle = _wholeTitle.Document.Range(originalStart, lefts[2] - 2);
        }

        private void InitializeFields()
        {
            IList<IField> fields = new List<IField>();
            IConfigurationItem config = _configuration.GetWorkItemConfigurationExtended(WorkItemType, fieldName =>
            {
                if (string.IsNullOrEmpty(fieldName))
                    return null;
                var fieldToEvaluate = Fields[fieldName];
                return fieldToEvaluate == null ? null : fieldToEvaluate.Value;
            });
            if (null != config)
            {
                foreach (IConfigurationFieldItem fieldItem in config.FieldConfigurations)
                {
                    // We need to map only configured items in IConfigurationFieldToLinkedItem
                    foreach (IConfigurationFieldAssignment linkedFieldItem in ConfigurationToLinkedItem.FieldAssignmentConfiguration)
                    {
                        if (linkedFieldItem.ReferenceName == fieldItem.ReferenceFieldName)
                        {
                            IField field = null;
                            Range content = null;
                            if (linkedFieldItem.FieldPosition == FieldPositionType.Hidden)
                            {
                                // ToDo: Finish this implementation
                                field = new HiddenField(fieldItem, config.GetConverter(fieldItem.ReferenceFieldName));
                            }
                            if (linkedFieldItem.FieldPosition == FieldPositionType.NumberedListItemFirstAddition)
                            {
                                content = NumberedListItemFirstAddition;
                            }
                            else if (linkedFieldItem.FieldPosition == FieldPositionType.NumberedListItemSecondAddition)
                            {
                                content = NumberedListItemSecondAddition;
                            }
                            else if (linkedFieldItem.FieldPosition == FieldPositionType.NumberedListItemThirdAddition)
                            {
                                content = NumberedListItemThirdAddition;
                            }
                            else if (linkedFieldItem.FieldPosition == FieldPositionType.NumberedListItem)
                            {
                                content = WorkItemTitle;
                            }
                            else if (linkedFieldItem.FieldPosition == FieldPositionType.Remainder)
                            {
                                content = WorkItemDescription;
                            }
                            if (content != null)
                            {
                                var newField = new WordTableField(content, fieldItem, config.GetConverter(fieldItem.ReferenceFieldName), false);
                                fields.Add(newField);
                            }
                            else if (field != null)
                                fields.Add(field);
                        }
                    }
                }
            }

            _fields = new WordTableFieldCollection(fields);
        }
        #endregion Private methods

        #region Implementation of IWorkItem
        /// <summary>
        /// Gets or sets the work item id of the current work item. This value has to be unique among all work items
        /// </summary>
        public int Id
        {
            get
            {
                if (Fields.Contains(Resources.ReferenceNameId))
                {
                    // check whether it is a new work item
                    if (string.IsNullOrEmpty(Fields[Resources.ReferenceNameId].Value) ||
                        Fields[Resources.ReferenceNameId].Value.Equals("*"))
                    {
                        return 0;
                    }
                    int value;

                    // seems to be an existing one. pare the int value to return the id
                    if (int.TryParse(Fields[Resources.ReferenceNameId].Value, out value))
                    {
                        return value;
                    }
                }
                return -1;
            }
            private set
            {
                Fields[Resources.ReferenceNameId].Value = value == 0
                                                              ? string.Empty
                                                              : value.ToString(CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Gets the work item revision.
        /// </summary>
        public int Revision
        {
            get
            {
                if (Fields.Contains(CoreFieldReferenceNames.Rev))
                {
                    int value;
                    if (int.TryParse(Fields[CoreFieldReferenceNames.Rev].Value, out value))
                    {
                        return value;
                    }
                }

                return 0;
            }
            private set
            {
                Fields[Resources.ReferenceNameRevision].Value = value == 0
                                                              ? string.Empty
                                                              : value.ToString(CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Gets work item title.
        /// </summary>
        public string Title
        {
            get
            {
                if (Fields.Contains(FieldReferenceNames.SystemTitle))
                {
                    return Fields[FieldReferenceNames.SystemTitle].Value;
                }
                return string.Empty;
            }
        }

        /// <summary>
        /// Gets the type of the work item
        /// </summary>
        public string WorkItemType { get; private set; }

        /// <summary>
        /// Gets all fields which are supported by the current work item
        /// </summary>
        public IFieldCollection Fields
        {
            get
            {
                if (null == _fields)
                {
                    var dispatcher = SyncServiceFactory.GetService<Dispatcher>();
                    if (dispatcher != null)
                    {
                        if (!dispatcher.CheckAccess())
                        {
                            var action = new Action(InitializeFields);
                            dispatcher.Invoke(action, null);
                        }
                        else
                        {
                            InitializeFields();
                        }
                    }
                    else
                    {
                        InitializeFields();
                    }

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
                // default since we dont care
                return false;
            }
        }

        /// <summary>
        /// Returns whether this work item is new
        /// </summary>
        public bool IsNew
        {
            get
            {
                return Id <= 0;
            }
        }

        /// <summary>
        /// Ids of linked work items, grouped by link type
        /// </summary>
        public Dictionary<IConfigurationLinkItem, int[]> Links
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// Creates a bookmark for this work item.
        /// </summary>
        public void CreateBookmark()
        {
            // Nothing to do. We don't support this feature.
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
        /// Adds the link.
        /// </summary>
        /// <param name="adapter">Adapter from which to look up missing references </param>
        /// <param name="ids">The id array.</param>
        /// <param name="linkType">The link type.</param>
        /// <param name="overwrite">if set to <c>true</c> overwrites any existing links of that type.</param>
        /// <returns></returns>
        public bool AddLinks(IWorkItemSyncAdapter adapter, int[] ids, string linkType, bool overwrite)
        {
            // Nothing to do. We don't support this feature.
            return true;
        }

        /// <summary>
        /// Adds the link.
        /// </summary>
        /// <param name="adapter">Adapter from which to look up missing references </param>
        /// <param name="ids">The id array.</param>
        /// <param name="linkType">The link type.</param>
        /// <param name="overwrite">if set to <c>true</c> overwrites any existing links of that type.</param>
        /// <returns></returns>
        public bool AddLinks(IWorkItemSyncAdapter adapter, int[] ids, string linkType, string linkedWorkItemTypes, bool overwrite)
        {
            // Nothing to do. We don't support this feature.
            return true;
        }

        /// <summary>
        /// Refresh work item to get latest state.
        /// </summary>
        public void Refresh()
        {
            // Nothing to do. We don't support this feature.
        }

        /// <summary>
        /// Get revision number for this work item. If the referenceFieldName is defined then:
        /// TFS: Gets revision number which belongs to the field -> last revision number for certain field.
        /// Word: Gets revision number as value from reference field name, if parameter not defined then use System.Rev.
        /// </summary>
        /// <param name="referenceFieldName">Reference field name.</param>
        /// <returns></returns>
        public int GetFieldRevision(string referenceFieldName)
        {
            var referenceName = referenceFieldName;
            if (string.IsNullOrEmpty(referenceName))
            {
                referenceName = FieldReferenceNames.SystemRev;
            }

            var revField = Fields.FirstOrDefault(field => field.ReferenceName.Equals(referenceName));
            if (revField != null)
            {
                int val;
                if (int.TryParse(revField.Value, out val))
                {
                    return val;
                }
            }

            return 0;
        }

        /// <summary>
        /// Gets a list of fields and their values for a given revision of the work item.
        /// </summary>
        /// <param name="revision">Revision for which to retrieve values.</param>
        /// <returns>
        /// List of fields for this revision.
        /// </returns>
        /// <exception cref="NotSupportedException"></exception>
        public IFieldCollection GetWorkItemByRevision(int revision)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not supported.
        /// </summary>
        /// <exception cref="NotSupportedException"></exception>
        public IConfigurationItem Configuration { get { throw new NotSupportedException();} }

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
            Id = 0;
            Revision = 0;
        }

        #endregion Implementation of IWorkItem
    }
}
