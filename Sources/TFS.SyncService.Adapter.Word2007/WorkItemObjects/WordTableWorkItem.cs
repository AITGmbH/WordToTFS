using AIT.TFS.SyncService.Adapter.Word2007.Properties;
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Exceptions;
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

namespace AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects
{
    /// <summary>
    /// Class implements <see cref="IWorkItem"/> and used in WORD Adapter.
    /// </summary>
    public class WordTableWorkItem : IWorkItemLinkedItems
    {
        #region Fields

        private readonly IConfigurationItem _configurationItem;
        private readonly IConfiguration _configuration;
        private readonly Dictionary<IConfigurationLinkItem, Range> _linkRangeCache = new Dictionary<IConfigurationLinkItem, Range>();

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTableWorkItem"/> class.
        /// </summary>
        /// <param name="table">Table to wrap.</param>
        /// <param name="workItemType">Work item to wrap.</param>
        /// <param name="configuration">Configuration of all work items</param>
        /// <param name="configurationItem">The configuration of this work item</param>
        /// <exception cref="ConfigurationException">Thrown if the configuration contains invalid cell column or row for standard fields.</exception>
        internal WordTableWorkItem(Table table, string workItemType, IConfiguration configuration, IConfigurationItem configurationItem)
        {
            if (table == null) throw new ArgumentNullException("table");
            if (workItemType == null) throw new ArgumentNullException("workItemType");
            if (configurationItem == null) throw new ArgumentNullException("configurationItem");

            Table = table;
            TableStart = table.Range.Start;
            WorkItemType = workItemType;
            _configuration = configuration;
            _configurationItem = configurationItem.Clone();

            // Create fields
            IList<IField> fields = new List<IField>();
            foreach (var fieldItem in _configurationItem.FieldConfigurations)
            {
                if (fieldItem.IsMapped)
                {
                    var cellRange = WordSyncHelper.GetCellRange(Table, fieldItem.RowIndex, fieldItem.ColIndex);

                    if (fieldItem.FieldValueType == FieldValueType.BasedOnVariable || fieldItem.FieldValueType == FieldValueType.BasedOnSystemVariable)
                    {
                      fields.Add(new StaticValueField(cellRange, fieldItem, true));
                    }
                    else
                    {
                      fields.Add(new WordTableField(cellRange, fieldItem, _configurationItem.GetConverter(fieldItem.ReferenceFieldName), true));
                    }
                }
                else
                {
                    // Initialize unmapped fields with default values
                    var hiddenField = new HiddenField(fieldItem, _configurationItem.GetConverter(fieldItem.ReferenceFieldName));
                    if (fieldItem.DefaultValue != null)
                    {
                        hiddenField.Value = fieldItem.DefaultValue.DefaultValue;
                    }

                    fields.Add(hiddenField);
                }
            }

            Fields = new WordTableFieldCollection(fields);

            IsNew = Id <= 0;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Gets the start of the table. This is cached for better performance of EqualsTable
        /// </summary>
        public int TableStart
        {
            get;
            private set;
        }

        /// <summary>
        /// The table mapped by this work item
        /// </summary>
        public Table Table { get; private set; }

        /// <summary>
        /// The method creates work items dependent on this work item
        /// </summary>
        /// <returns>Collection of dependent work items.</returns>
        /// <exception cref="ConfigurationException">Thrown if the configuration contains invalid cell column or row for FieldToLinkedItemConfiguration.</exception>
        public IWorkItemCollection CreateLinkedWorkItems()
        {
            var retValue = new WordTableWorkItemCollection();
            LinkedWorkItems = new Dictionary<string, IWorkItemCollection>();
            // Get the associated configuration

            if (!_configurationItem.ConfigurationSupportsLinkedItem)
            {
                return retValue;
            }
            if (_configurationItem.FieldToLinkedItemConfiguration == null)
            {
                return retValue;
            }
            if (_configurationItem.FieldToLinkedItemConfiguration.Count == 0)
            {
                return retValue;
            }
            foreach (var configurationLinkedItem in _configurationItem.FieldToLinkedItemConfiguration)
            {
                // Get the cell of the table, where the linked work items are defined.
                var cellRange = WordSyncHelper.GetCellRange(Table, configurationLinkedItem.RowIndex, configurationLinkedItem.ColIndex);
                if (cellRange == null)
                {
                    continue;
                }
                // At this moment only 'NumberedList' supported
                if (configurationLinkedItem.WorkItemBindType != WorkItemBindType.NumberedList)
                {
                    continue;
                }
                var numberedListItems = new List<Range>();
                foreach (Paragraph paragraph in cellRange.Paragraphs)
                {
                    if (paragraph.Range.ListFormat != null && paragraph.Range.ListFormat.ListLevelNumber == 1 &&
                        paragraph.Range.ListFormat.ListValue > 0)
                    {
                        // We need to store only ranges of 'valid' list items
                        // Valid list item is for us item with level 1 and value bigger as 0
                        numberedListItems.Add(paragraph.Range);
                    }
                }
                for (int index = 0; index < numberedListItems.Count; index++)
                {
                    var title = numberedListItems[index];
                    int end = cellRange.End - 1; // -1 to cut the last charactar in the cell - \v
                    if (index + 1 < numberedListItems.Count) end = numberedListItems[index + 1].Start;
                    var description = title.Document.Range(title.End, end);
                    var newWorkItem = new WordNumberedListItemWorkItem(configurationLinkedItem, title, description,
                                                                       _configuration);
                    retValue.Add(newWorkItem);
                    if (!LinkedWorkItems.ContainsKey(configurationLinkedItem.LinkType.ToString()))
                    {
                        LinkedWorkItems[configurationLinkedItem.LinkType.ToString()] = new WordTableWorkItemCollection();
                    }
                    LinkedWorkItems[configurationLinkedItem.LinkType.ToString()].Add(newWorkItem);
                }
            }
            return retValue;
        }
        
        #endregion Public methods

        #region IWorkItem Members

        /// <summary>
        /// Gets a configuration for this work item and all its field. This also takes into account fields that
        /// are added due to headers before this work item in the document.
        /// </summary>
        public IConfigurationItem Configuration
        {
            get
            {
                _configurationItem.FieldConfigurations.Clear();

                foreach (var field in Fields)
                {
                    _configurationItem.FieldConfigurations.Add(field.Configuration.Clone());
                }

                return _configurationItem;
            }
        }

        /// <summary>
        /// Determines the type of the work item
        /// </summary>
        public string WorkItemType { get; private set; }

        /// <summary>
        /// The work item id of the current work item. This value has to be unique among all work items.
        /// </summary>
        public int Id
        {
            get
            {
                if (Fields.Contains(Resources.ReferenceNameId))
                {
                    int value;
                    if (int.TryParse(Fields[Resources.ReferenceNameId].Value, out value))
                    {
                        return value;
                    }
                }
                return 0;
            }

            private set
            {
                // Set it to empty, as fix for "Delete IDs" option
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
        /// All fields which are supported by the current work item
        /// </summary>
        public IFieldCollection Fields { get; private set; }

        /// <summary>
        /// Returns whether this work item has been modified
        /// </summary>
        public bool IsDirty
        {
            get
            {
                // no checking for word items
                return true;
            }
        }

        /// <summary>
        /// Returns whether this work item is new
        /// </summary>
        public bool IsNew
        {
            get;
            private set;
        }

        /// <summary>
        /// Ids of linked work items, grouped by link type
        /// </summary>
        /// <exception cref="ConfigurationException">Thrown if the configuration contains invalid cell column or row for configured links.</exception>
        public Dictionary<IConfigurationLinkItem, int[]> Links
        {
            get
            {
                var links = new Dictionary<IConfigurationLinkItem, int[]>();
                if (_configurationItem == null)
                {
                    return links;
                }

                // For each link configuration, try to parse mapped cell contents to ids
                foreach (var linkItem in _configurationItem.Links)
                {
                    var cellRange = GetLinkRange(linkItem);
                    if (cellRange == null)
                    {
                        continue;
                    }

                    // Split cell text by defined link separator
                    var text = GetCellText(cellRange);
                    var formattedLinks = text.Split(new[] { linkItem.LinkSeparator }, StringSplitOptions.RemoveEmptyEntries);

                    var ids = new List<int>();

                    // Try to retrieve the id from each formatted link)))
                    foreach (var formattedLink in formattedLinks)
                    {
                        var id = linkItem.GetWorkItemId(formattedLink);
                        if (id.HasValue)
                        {
                            ids.Add(id.Value);
                        }
                    }

                    SyncServiceTrace.D(Resources.LogService_AccessToLinkMessage, Id, linkItem.LinkValueType,
                                       text,
                                       string.Join(",",
                                                   ids.Select(x => x.ToString(CultureInfo.InvariantCulture)).ToArray()));
                    links.Add(linkItem, ids.ToArray());
                }

                return links;
            }
        }

        /// <summary>
        /// Adds a single link to another work item
        /// </summary>
        /// <param name="adapter">Adapter used to look up more information about the referenced work items.</param>
        /// <param name="id">The id of the other work item.</param>
        /// <param name="link">Link type.</param>
        private void AddLink(IWorkItemSyncAdapter adapter, int id, IConfigurationLinkItem link)
        {
            var range = GetLinkRange(link);
            if (range == null)
            {
                return;
            }

            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            var bookmarks = Table.Range.Document.Bookmarks;
            var workItem = adapter.WorkItems.Find(id);

            // if we can look up the work item, format the output. Otherwise only display the work item id
            var linkTitle = workItem == null ? id.ToString(CultureInfo.InvariantCulture) : link.Format(workItem);

            // insert hyper link to target work item bookmark
            range.Text = linkTitle;
            if (bookmarks.Exists(id))
            {
                range.Hyperlinks.Add(range, SubAddress: bookmarks.Item(id));
                range.SetRange(range.End + 1, range.End + 1);
            }

            range.Collapse(WdCollapseDirection.wdCollapseEnd);
        }

        /// <summary>
        /// Adds multiple links to another work items to this work item.
        /// </summary>
        /// <param name="adapter">Adapter used to look up more information about the referenced work items.</param>
        /// <param name="ids">The ids of the other work items.</param>
        /// <param name="linkType">The reference name of the link type.</param>
        /// <param name="overwrite">Not used. Always overwrites cell content</param>
        /// <returns>True when the link was fully added, false otherwise.</returns>
        /// <exception cref="ConfigurationException">Thrown if the configuration contains invalid cell column or row for link configuration.</exception>
        public bool AddLinks(IWorkItemSyncAdapter adapter, int[] ids, string linkType, bool overwrite)
        {
            return AddLinks(adapter, ids, linkType, string.Empty, overwrite);
        }

        /// <summary>
        /// Adds multiple links to another work items to this work item.
        /// </summary>
        /// <param name="adapter">Adapter used to look up more information about the referenced work items.</param>
        /// <param name="ids">The ids of the other work items.</param>
        /// <param name="linkType">The reference name of the link type.</param>
        /// <param name="linkedWorkItemTypes">Determines a comma separated list of Work Item Types that can be used to filter the links in addition to the linkValueType.</param>
        /// <param name="overwrite">Not used. Always overwrites cell content</param>
        /// <returns>True when the link was fully added, false otherwise.</returns>
        /// <exception cref="ConfigurationException">Thrown if the configuration contains invalid cell column or row for link configuration.</exception>
        public bool AddLinks(IWorkItemSyncAdapter adapter, int[] ids, string linkType, string linkedWorkItemTypes, bool overwrite)
        {
            if (_configurationItem == null)
            {
                return false;
            }

            var link = string.IsNullOrEmpty(linkedWorkItemTypes)
                ? _configurationItem.Links.FirstOrDefault(x => x.LinkValueType.Equals(linkType))
                : _configurationItem.Links.FirstOrDefault(x => x.LinkValueType.Equals(linkType) && string.Equals(x.LinkedWorkItemTypes, linkedWorkItemTypes));

            if (link == null)
            {
                return false;
            }

            var range = GetLinkRange(link);
            if (range == null)
            {
                return false;
            }

            range.Delete();
            var rangeStart = range.Start;

            // Add links, delete contents of cell before the first item.
            if (ids != null)
            {
                var firstPass = true;
                foreach (var id in ids)
                {
                    if (firstPass == false)
                    {
                        range.InsertAfter(link.LinkSeparator);
                    }

                    firstPass = false;
                    AddLink(adapter, id, link);
                }
            }

            range.Start = rangeStart;
            return true;
        }

        /// <summary>
        /// Adds the automatically created bookmark as well as the bookmarks defined in the config file.
        /// </summary>
        public void CreateBookmark()
        {
            var bookmarkName = string.Concat("w2t", Id);
            if (!Table.Range.Bookmarks.Exists(bookmarkName))
            {
                Table.Cell(Configuration.ReqTableCellRow, Configuration.ReqTableCellCol).Range.Bookmarks.Add(
                    bookmarkName);
            }
        }

        /// <summary>
        /// Return whether this work items had errors during the validation on the word side
        /// </summary>
        public bool HasWordValidationError
        {
            get;
            set;
        }

        /// <summary>
        /// Word table work items cannot be refreshed yet. The state of the fields is not supposed to changed
        /// during the lifetime of a word adapter. Create a new adapter to reread all tables.
        /// </summary>
        public void Refresh()
        {
            // You could implement this by invalidating all field caches.
        }

        /// <summary>
        /// Get revision number for this work item. If the referenceFieldName is defined then
        /// gets revision number as value from reference field name, if parameter not defined then use System.Rev.
        /// </summary>
        /// <param name="referenceFieldName">Reference field name.</param>
        /// <returns></returns>
        public int GetFieldRevision(string referenceFieldName)
        {
            Guard.ThrowOnArgumentNull(referenceFieldName, "referenceFieldName");
            return Fields.Contains(referenceFieldName) ? Revision : 0;
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
        /// Returns whether this and another work item are wrapper for the same work item.
        /// </summary>
        /// <param name="other">Other work item.</param>
        /// <returns>
        /// True, if both work item wrappers are wrapper for the same work item.
        /// </returns>
        public bool IsSameWorkItem(IWorkItem other)
        {
            if (other == this)
            {
                return true;
            }

            var otherWord = other as WordTableWorkItem;
            if (otherWord != null)
            {
                return otherWord.TableStart == TableStart;
            }

            return false;
        }

        /// <summary>
        /// Delete Id and Revision
        /// </summary>
        public void DeleteIdAndRevision()
        {
            Id = 0;
            Revision = 0;
        }

        #endregion

        #region Implementation of IWorkItemLinkedItems

        /// <summary>
        /// The property gets the Dictionary of all linked items ordered in 'link type' groups.
        /// Key of dictionary means the link type and the values are all work to link with mentioned link type.
        /// </summary>
        public IDictionary<string, IWorkItemCollection> LinkedWorkItems { get; private set; }

        #endregion Implementation of IWorkItemLinkedItems

        #region Private methods

        /// <summary>
        /// Wraps cached access to the ranges of link fields.
        /// </summary>
        private Range GetLinkRange(IConfigurationLinkItem linkConfiguration)
        {
            if (_linkRangeCache.ContainsKey(linkConfiguration) == false)
            {
                var range = WordSyncHelper.GetCellRange(Table, linkConfiguration.RowIndex, linkConfiguration.ColIndex);
                _linkRangeCache.Add(linkConfiguration, range);
            }

            return _linkRangeCache[linkConfiguration];
        }

        /// <summary>
        /// Gets the text of a cell without the end of cell marker.
        /// </summary>
        /// <param name="range">Range of the cell</param>
        /// <returns>Text of the range without last character (end of cell marker)</returns>
        private static string GetCellText(Range range)
        {
            if (range.Text == null)
            {
                return string.Empty;
            }

            return range.Text.TrimEnd(new[] { '\a', '\r' });
        }

        #endregion
    }
}