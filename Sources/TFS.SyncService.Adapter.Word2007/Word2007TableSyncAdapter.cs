using AIT.TFS.SyncService.Adapter.Word2007.Properties;
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemCollections;
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using Microsoft.Office.Interop.Word;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace AIT.TFS.SyncService.Adapter.Word2007
{
    /// <summary>
    /// Implementation of word 2007 sync adapter - table version.
    /// </summary>
    internal class Word2007TableSyncAdapter : Word2007SyncAdapter
    {
        private readonly Stack<WordTableWorkItem> _headers = new Stack<WordTableWorkItem>();

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Word2007TableSyncAdapter"/> class.
        /// </summary>
        /// <param name="configuration">Configuration to work with</param>
        /// <param name="document">Word document to work with.</param>
        public Word2007TableSyncAdapter(Document document, IConfiguration configuration)
            : base(document, configuration)
        {
        }
        #endregion Constructors

        #region Overridden methods
        /// <summary>
        /// Table adapter specific load. Loads <c>WordTableWorkItems</c> from tables.
        /// </summary>
        /// <param name="ids">A List of ids of the work items to load</param>
        protected override bool LoadStructure(int[] ids)
        {
            return LoadStructure((table, item) => LoadById(table, item, ids));
        }

        /// <summary>
        /// Table adapter specific load. Loads <c>WordTableWorkItems</c> from tables.
        /// </summary>
        /// <param name="loadCriteria">Selects which items to actually load. This is a performance enhancer...</param>
        protected bool LoadStructure(Func<Table, IConfigurationItem, bool> loadCriteria)
        {
            WorkItems = new WordTableWorkItemCollection();

            if (Document.Tables == null) return true;

            foreach (Table table in Document.Tables)
            {
                try
                {
                    var workItem = CreateWorkItem(table, loadCriteria);

                    // add successfully created items if the id is in the list or no list was set
                    if (workItem != null)
                    {
                        ApplyHeaders(workItem);
                        WorkItems.Add(workItem);

                        // Load embedded work items
                        foreach (var subItem in workItem.CreateLinkedWorkItems())
                        {
                            ApplyHeaders(workItem);
                            WorkItems.Add(subItem);
                        }
                    }
                    else
                    {
                        var header = CreateHeader(table);

                        // If the table was a header, remove all headers of higher or the same level from the stack
                        if (header != null)
                        {
                            while (_headers.Count > 0 && _headers.Peek().Configuration.Level >= header.Configuration.Level)
                            {
                                _headers.Pop();
                            }
                            _headers.Push(header);
                            SyncServiceTrace.D("Found new header, new stack: {0}", string.Join(",", _headers.Select(x => x.Configuration.WorkItemType).ToArray()));
                        }
                    }
                }
                catch (Exception ex)
                {
                    SyncServiceTrace.LogException(ex);
                    var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
                    if (infoStorage != null)
                    {
                        infoStorage.NotifyError(Resources.Error_LoadWordWI, ex);
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Gets the work items represented by the currently selected tables.
        /// </summary>
        /// <returns>The selected work items or an empty enumeration if no items are selected.</returns>
        public override IEnumerable<IWorkItem> GetSelectedWorkItems()
        {

            if (!IsOpen || Document.ActiveWindow == null)
            {
                return Enumerable.Empty<IWorkItem>();
            }
            var tables = Document.ActiveWindow.Selection.Tables;
            if (tables == null)
            {
                return Enumerable.Empty<IWorkItem>();
            }
            var selectAll = new Func<Table, IConfigurationItem, bool>((table, item) => true);

            return from Table t in tables where CreateWorkItem(t, selectAll ) != null select (IWorkItem)CreateWorkItem(t, selectAll);
        }

        /// <summary>
        /// Gets the header that is currently selected
        /// </summary>
        /// <returns></returns>
        public override IEnumerable<IWorkItem> GetSelectedHeader()
        {
            if (!IsOpen || Document.ActiveWindow == null)
            {
                return Enumerable.Empty<IWorkItem>();
            }
            var tables = Document.ActiveWindow.Selection.Tables;
            if (tables == null)
            {
                return Enumerable.Empty<IWorkItem>();
            }

            return from Table t in tables
                   where CreateHeader(t) != null
                   select (IWorkItem) CreateHeader(t);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Parses a <c>Table</c> into a <c>IWorkItem</c>. Uses the currently set configuration to determine
        /// what work item types are available.
        /// </summary>
        /// <param name="table">table representing the work item</param>
        /// <param name="loadCriteria">Selects which items to load before actually creating the wrapper. Performance enhancer.</param>
        /// <returns>A work item represented by the table or <c>null</c> if the type
        /// of the work item could not be determined </returns>
        private WordTableWorkItem CreateWorkItem(Table table, Func<Table, IConfigurationItem, bool> loadCriteria)
        {
            var config = GetConfigurationItemFromTable(table, Configuration.GetConfigurationItems());

            if (config == null)
            {
                SyncServiceTrace.W(Resources.LogService_TableToWorkItemException);
                return null;
            }

            if(loadCriteria(table, config) == false)
            {
                return null;
            }

            return new WordTableWorkItem(table, config.WorkItemType, Configuration, config);
        }

        /// <summary>
        /// Peek at the id in a table and only create wrapper if id is in given set
        /// </summary>
        private static bool LoadById(Table table, IConfigurationItem configuration, int[] ids)
        {
            if (ids == null)
            {
                return true;
            }

            var fieldConfiguration = configuration.FieldConfigurations.FirstOrDefault(x => x.ReferenceFieldName == CoreFieldReferenceNames.Id);
            if (fieldConfiguration == null)
            {
                return false;
            }

            var range = WordSyncHelper.GetCellRange(table, fieldConfiguration.RowIndex, fieldConfiguration.ColIndex);
            int id;
            if (range == null || int.TryParse(range.Text.Replace("\r\a", string.Empty), out id) == false)
            {
                return false;
            }

            return ids.Contains(id);
        }

        /// <summary>
        /// Parses a <c>Table</c> into a <c>IWorkItem</c>. Uses the currently set configuration to determine
        /// what work item types are available.
        /// </summary>
        /// <param name="table">table representing the work item</param>
        /// <returns>A work item represented by the table or <c>null</c> if the type
        /// of the work item could not be determined </returns>
        private WordTableWorkItem CreateHeader(Table table)
        {
            IConfigurationItem config = GetConfigurationItemFromTable(table, Configuration.Headers);

            if (config == null)
            {
                SyncServiceTrace.W(Resources.LogService_TableToWorkItemException);
                return null;
            }

            return new WordTableWorkItem(table, config.WorkItemType, Configuration, config);
        }

        /// <summary>
        /// Returns a <see cref="IConfigurationItem"/> that is represented by the <paramref name="table"/>
        /// </summary>
        /// <param name="table">Table representation of a <see cref="IConfigurationItem"/></param>
        /// <param name="definitions">The definitions in which to look for an appropriate configuration </param>
        /// <returns><see cref="IConfigurationItem"/> of the table or null if none is found</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", MessageId = "Body")]
        private static IConfigurationItem GetConfigurationItemFromTable(Table table, IEnumerable<IConfigurationItem> definitions)
        {
            if (table == null) throw new ArgumentNullException("table");

            var tableRowsCountCache = table.Rows.Count;
            var tableColumnsCountCache = table.Columns.Count;
            var tableCellRangesCache = new Range[tableRowsCountCache,tableColumnsCountCache];

            IConfigurationItem exactConfiguration = null;
            IConfigurationItem regexConfiguration = null;

            foreach (var configuration in definitions)
            {
                // Get the cell range.
                if (tableRowsCountCache < configuration.ReqTableCellRow || tableColumnsCountCache < configuration.ReqTableCellCol)
                {
                    continue;
                }


                var range = tableCellRangesCache[configuration.ReqTableCellRow - 1, configuration.ReqTableCellCol - 1];
                if (range == null)
                {
                    try
                    {
                        range = WordSyncHelper.GetCellRange(table, configuration.ReqTableCellRow, configuration.ReqTableCellCol);
                        tableCellRangesCache[configuration.ReqTableCellRow - 1, configuration.ReqTableCellCol - 1] = range;
                    }
                    catch (ConfigurationException)
                    {
                        SyncServiceTrace.I(Resources.TheReferencesTableIdentifierCellNotExists, configuration.ReqTableCellRow, configuration.ReqTableCellCol);
                        continue;
                    }
                }

                if (range == null)
                {
                    continue;
                }

                // Match cell text with exact work item type mapping
                var cellText = WordSyncHelper.FormatCellContent(range.Text);
                if (configuration.WorkItemTypeMapping.Equals(cellText))
                {
                    exactConfiguration = configuration;
                }
                // Compare the regex expression for a direct mathc
                if (configuration.ReqTableIdentifierExpression.Equals(cellText))
                {
                    exactConfiguration = configuration;
                }

                // Check regular expression.
                if (!string.IsNullOrEmpty(configuration.ReqTableIdentifierExpression)
                    && Regex.IsMatch(cellText, configuration.ReqTableIdentifierExpression))
                {
                    regexConfiguration = configuration;
                }
            }

            //If a exact match is available return it, else the regex
            if (exactConfiguration != null)
            {
                return exactConfiguration;

            }
            else if (regexConfiguration != null)
            {
                return regexConfiguration;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Applies all active headers to a work item. If a header contains fields not mapped by the work item,
        /// it will add these fields as configured in the header. If the work item already contains such a field,
        /// it is replaced with a split field, that reads from the header but writes to the work item. If the
        /// work item does not contain a field, a SplitField is added that writes to a HiddenField.
        /// </summary>
        /// <param name="item">Work item to which to apply headers</param>
        private void ApplyHeaders(IWorkItem item)
        {
            var appliedFields = new List<string>();
            foreach(var header in _headers)
            {
                foreach (var field in header.Fields)
                {
                    if (appliedFields.Contains(field.ReferenceName)) continue;
                    //Fix for Bug 17650
                    var converter = item.Configuration.GetConverter(field.ReferenceName);
                    IField writeField = new HiddenField(field.Configuration, converter);
                    if (item.Fields.Contains(field.ReferenceName))
                    {
                        writeField = item.Fields[field.ReferenceName];
                        item.Fields.Remove(item.Fields[field.ReferenceName]);
                    }

                    appliedFields.Add(field.ReferenceName);
                    item.Fields.Add(new SplitField(item, field, writeField));
                }
            }
        }

        #endregion
    }
}