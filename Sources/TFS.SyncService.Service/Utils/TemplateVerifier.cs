using AIT.TFS.SyncService.Contracts;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Factory;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace AIT.TFS.SyncService.Service.Utils
{
    /// <summary>
    /// Checks a template for field names that do not exist on the server.
    /// </summary>
    public class TemplateVerifier
    {
        private readonly IList<IConfigurationFieldItem> _missingFields = new Collection<IConfigurationFieldItem>();
        private readonly IList<IConfigurationFieldItem> _wrongMappedFields = new Collection<IConfigurationFieldItem>();

        /// <summary>
        /// Gets the missing fields.
        /// </summary>
        /// <value>The missing fields.</value>
        public IList<IConfigurationFieldItem> MissingFields
        {
            get
            {
                return _missingFields;
            }
        }


        /// <summary>
        /// Gets the wrong mapped fields.
        /// </summary>
        /// <value>The wrong mapped fields.</value>
        public IList<IConfigurationFieldItem> WrongMappedFields
        {
            get
            {
                return _wrongMappedFields;
            }
        }

        /// <summary>
        /// Verifies the template mapping.
        /// </summary>
        /// <param name="serviceModel">The service document model.</param>
        /// <param name="syncAdapter">The sync adapter.</param>
        /// <param name="configuration">Configuration with the field to check for existence on the server</param>
        public bool VerifyTemplateMapping(ISyncServiceDocumentModel serviceModel, ITfsService syncAdapter, IConfiguration configuration)
        {
            Guard.ThrowOnArgumentNull(serviceModel, "serviceModel");
            Guard.ThrowOnArgumentNull(syncAdapter, "syncAdapter");
            Guard.ThrowOnArgumentNull(configuration, "configuration");

            _missingFields.Clear();

            // check each mapping field
            foreach (var configurationItem in configuration.GetConfigurationItems())
            {
                if (!configurationItem.RelatedSchema.Equals(serviceModel.MappingShowName)) continue;
                foreach (var fieldItem in configurationItem.FieldConfigurations)
                {
                    var fieldExists = syncAdapter.FieldDefinitions.Cast<FieldDefinition>().Any(field => fieldItem.ReferenceFieldName.Equals(field.ReferenceName));

                    if (!fieldExists)
                    {
                        // add missing fields to collection if it is not based on variable type
                        fieldItem.SourceMappingName = configurationItem.WorkItemTypeMapping;

                        if (!fieldItem.FieldValueType.IsVariable())
                        {
                            _missingFields.Add(fieldItem);
                        }
                    }
                        //Field Exists, check if the defintions are right
                    else
                    {
                        //If the field on Word is handeled as HTML, check the counterpart on the tfs
                        if (fieldItem.FieldValueType == FieldValueType.HTML)
                        {
                            //Get the TFS Item
                            var tfsFieldItem = syncAdapter.FieldDefinitions[fieldItem.ReferenceFieldName];

                            if (tfsFieldItem.FieldType != FieldType.Html)
                            {
                                fieldItem.SourceMappingName = configurationItem.WorkItemTypeMapping;
                                _wrongMappedFields.Add(fieldItem);
                            }
                        }
                    }
                }
            }

            // return true if there were no missing fields or no wrong mapped field, else false.
            if (_missingFields.Count == 0 && _wrongMappedFields.Count == 0)
            {
                return true;
            }

            return false;
        }
    }
}