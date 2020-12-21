
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections.Generic;
using System.Linq;

namespace AIT.TFS.SyncService.Service.Utils
{
    /// <summary>
    /// Checks if fields are mapped right between a list of work items and the server.
    /// </summary>
    public class FieldVerifier
    {
        readonly Dictionary<WordTableWorkItem, List<WordTableField>> _wrongMappedFieldsDictionary = new Dictionary<WordTableWorkItem, List<WordTableField>>();


        /// <summary>
        /// The dictionary with the Workitems and the wrong mapped fields
        /// </summary>
        public Dictionary<WordTableWorkItem, List<WordTableField>> WrongMappedFieldsDictionary
        {
            get
            {
                return _wrongMappedFieldsDictionary;
            }
        }

        /// <summary>
        /// Verifies the template mapping.
        /// </summary>
        /// <param name="wordWorkItems">The work items that are published</param>
        /// <param name="destinationSyncAdapter">The sync adapter.</param>
        /// <param name="configuration">Configuration with the field to check for existence on the server</param>
        public bool VerifyTemplateMapping(IEnumerable<IWorkItem> wordWorkItems, IWorkItemSyncAdapter destinationSyncAdapter, IConfiguration configuration)
        {
            Guard.ThrowOnArgumentNull(wordWorkItems, "workItems");
            Guard.ThrowOnArgumentNull(destinationSyncAdapter, "destinationSyncAdapter");
            Guard.ThrowOnArgumentNull(configuration, "configuration");

            _wrongMappedFieldsDictionary.Clear();

            //Get the service
            var tfsService = destinationSyncAdapter as ITfsService;

            if (tfsService != null)
            {


                //Loop through all itenms
                foreach (WordTableWorkItem singleWorkItem in wordWorkItems)
                {
                    foreach (IField fieldItem in singleWorkItem.Fields)
                    {
                        if (fieldItem is WordTableField)
                        {
                        //Check if the field exists
                        var fieldExists = tfsService.FieldDefinitions.Cast<FieldDefinition>().Any(field => fieldItem.Configuration.ReferenceFieldName.Equals(field.ReferenceName));
                        if (fieldExists)
                        {
                            //If the field on Word is handeled as HTML, check the counterpart on the tfs
                            if (fieldItem.Configuration.FieldValueType == (FieldValueType.HTML))
                            {
                                //Get the TFS Item and check if it is HTML
                                var tfsFieldItem = tfsService.FieldDefinitions[fieldItem.Configuration.ReferenceFieldName];
                                if (tfsFieldItem.FieldType != FieldType.Html)
                                {
                                    //Add it to the list with files, this list is later used to set the flags, depending on the users decision
                                    List<WordTableField> fieldList;
                                    if (_wrongMappedFieldsDictionary.TryGetValue(singleWorkItem, out fieldList))
                                    {
                                        ((WordTableField)fieldItem).ParseHtmlAsPlaintext = true;
                                        fieldList.Add((WordTableField)fieldItem);
                                    }
                                    else
                                    {
                                        fieldList = new List<WordTableField>();
                                        ((WordTableField)fieldItem).ParseHtmlAsPlaintext = true;
                                        fieldList.Add((WordTableField)fieldItem);
                                        _wrongMappedFieldsDictionary.Add(singleWorkItem, fieldList);
                                    }
                                }
                            }
                        }
                        }
                    }
                }
            }
            // return true if there were no missing fields or no wrong mapped field, else false.
            return _wrongMappedFieldsDictionary.Count != 0;
        }
    }
}