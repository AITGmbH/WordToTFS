using AIT.TFS.SyncService.Contracts.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace AIT.TFS.SyncService.Contracts.Model
{
    /// <summary>
    /// Interface defines the functionality of the model for one document.
    /// </summary>
    public interface ISyncServiceDocumentModel : INotifyPropertyChanged
    {
        /// <summary>
        /// Gets or sets a value indicating whether an operation is running.
        /// </summary>
        bool OperationInProgress { get; set; }

        /// <summary>
        /// Gets the associated document.
        /// </summary>
        object WordDocument { get; }

        /// <summary>
        /// Gets the name of associated document. If associated document is not valid string.Empty returned.
        /// </summary>
        string DocumentName { get; }

        /// <summary>
        /// Gets the flag that tells the active document is bound to TFS project.
        /// </summary>
        bool TfsDocumentBound { get; }

        /// <summary>
        /// Gets the TFS server where active document is bound. See the <see cref="TfsDocumentBound"/>.
        /// </summary>
        string TfsServer { get; }

        /// <summary>
        /// Gets the TFS project where active document is bound. See the <see cref="TfsDocumentBound"/>.
        /// </summary>
        string TfsProject { get; }

        /// <summary>
        /// Gets the show name of mapping to use with active document.
        /// </summary>
        /// <remarks>
        /// The mapping show name encapsulates many information. For example mapping file, template file.
        /// </remarks>
        string MappingShowName { get; set; }

        /// <summary>
        /// Read query configuration properties from document.
        /// </summary>
        /// <param name="queryConfiguration">IQueryConfiguration instance.</param>
        void ReadQueryConfiguration(IQueryConfiguration queryConfiguration);

        /// <summary>
        /// Save query configuration properties to document.
        /// </summary>
        /// <param name="queryConfiguration">IQueryConfiguration instance.</param>
        void SaveQueryConfiguration(IQueryConfiguration queryConfiguration);

        /// <summary>
        /// Clears the default value for given field.
        /// </summary>
        /// <param name="fieldName">Name of the field</param>
        void ClearDefaultValue(string fieldName);

        /// <summary>
        /// Methods sets the desired default value for given field in active document.
        /// </summary>
        /// <param name="fieldName">Name of the field to store the default value for.</param>
        /// <param name="fieldDefaultValue">Default value to store for the given field.</param>
        /// <param name="workItemType">Type of the affected work item.</param>
        void SetFieldDefaultValue(string fieldName, string fieldDefaultValue, string workItemType);

        /// <summary>
        /// Method gets the stored default value for given field.
        /// </summary>
        /// <param name="fieldName">Name of the field to get the default value for.</param>
        /// <param name="workItemType">Work item type of the affected work item.</param>
        /// <returns>Default value. Null if not defined.</returns>
        string GetFieldDefaultValue(string fieldName, string workItemType);

        /// <summary>
        /// Call this method to open the dialog for picking tfs servers and project. 
        /// this will bind the active document to the TFS project.
        /// </summary>
        void BindProject();

        /// <summary>
        /// Call this method to bind the active document to the given TFS project
        /// </summary>
        /// <param name="tfsServer"></param>
        /// <param name="tfsProject"></param>
        void BindProject(string tfsServer, string tfsProject);

        /// <summary>
        /// Call this method to unbound the active document from the TFS project.
        /// </summary>
        void UnbindProject();

        /// <summary>
        /// Saves data to document
        /// </summary>
        void Save();

        /// <summary>
        /// Gets a fields that are missing on the server for the current selected template.
        /// </summary>
        IList<IConfigurationFieldItem> GetMissingFields();

        /// <summary>
        /// Gets the fields that are mapped wrong on the current template
        /// </summary>
        IList<IConfigurationFieldItem> GetWronglyMappedFields();

        /// <summary>
        /// Gets or sets into document Area or Iteration path.
        /// </summary>
        string AreaIterationPath { get; set; }

        /// <summary>
        /// Gets the configuration for the associated document
        /// </summary>
        IConfiguration Configuration { get; }


        /// <summary>
        /// Occurs when variable changed in any document.
        /// </summary>
        event EventHandler<SyncServiceEventArgs> DocumentVariableChanged;

        #region Test report related

        /// <summary>
        /// Occurs if any data in test report area was changed.
        /// </summary>
        event EventHandler<SyncServiceEventArgs> TestReportDataChanged;

        /// <summary>
        /// Gets or sets the flag telling if the test report is running.
        /// </summary>
        bool TestReportRunning { get; set; }

        /// <summary>
        /// Gets or sets the flag telling if the document contains the test report - specification report or result report.
        /// </summary>
        /// <remarks>Value is stored in document variable.</remarks>
        bool TestReportGenerated { get; set; }

        /// <summary>
        /// The method reads the data associated with given key from document.
        /// </summary>
        /// <param name="key">Key to get the data for.</param>
        /// <returns>Read data. <c>null</c> if data was not written.</returns>
        string ReadTestReportData(string key);

        /// <summary>
        /// The method writes the given data with given associated key to document.
        /// </summary>
        /// <param name="key">Key to store the data for.</param>
        /// <param name="data">Data to store.</param>
        void WriteTestReportData(string key, string data);

        /// <summary>
        /// The method clears the data associated with given key from document.
        /// </summary>
        /// <param name="key">Key to clear the associated data.</param>
        void ClearTestReportData(string key);

        #endregion Test related
    }
}