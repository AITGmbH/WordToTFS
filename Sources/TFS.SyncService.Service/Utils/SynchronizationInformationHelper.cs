using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;

namespace AIT.TFS.SyncService.Service.Utils
{

    /// <summary>
    /// Helper to handle the synchronization between tfs and word
    /// </summary>
    public static class SynchronizationInformationHelper
    {
        /// <summary>
        /// Returns a list with all fields that have been changed in TFS compared to the current word revision and can be imported to Word.
        /// </summary>
        /// <param name="changedTfsFields">List with changed TFS fields.</param>
        /// <param name="tfsWorkItem">Work item on TFS</param>
        /// <returns>List of different fields.</returns>
        public static Collection<IField> GetRefreshableTfsFields(IList<IField> changedTfsFields, IWorkItem tfsWorkItem)
        {
            return new Collection<IField>(changedTfsFields.Where(
                x =>
                x.Configuration.Direction != Direction.PublishOnly &&
                (x.Configuration.Direction != Direction.GetOnly || tfsWorkItem.IsNew)).ToList());
        }

        /// <summary>
        /// Returns a list with all fields that have been changed in TFS compared to the current word revision.
        /// </summary>
        /// <param name="tfsWorkItem">Work item on TFS</param>
        /// <param name="wordWorkItem">Work item in word.</param>
        /// <param name="ignoreFormatting">Sets whether formatting is ignored.</param>
        /// <returns>List of different fields.</returns>
        public static Collection<IField> GetChangedTfsFields(IWorkItem wordWorkItem, IWorkItem tfsWorkItem, bool ignoreFormatting)
        {
            Guard.ThrowOnArgumentNull(wordWorkItem, "wordWorkItem");
            Guard.ThrowOnArgumentNull(tfsWorkItem, "tfsWorkItem");

            if (wordWorkItem.IsNew)
            {
                return new Collection<IField>();
            }

            //var originalTfsRevision = tfsWorkItem.GetRevision(wordWorkItem.Revision);
            var actuallyChangedFields = tfsWorkItem.Fields.Where(x => tfsWorkItem.GetFieldRevision(x.ReferenceName) > wordWorkItem.Revision).ToList();

            var actuallyChangedWordFields = GetFieldsWithDifferentValues(wordWorkItem.Fields.ToList(), actuallyChangedFields, ignoreFormatting).ToList();
            return new Collection<IField>(tfsWorkItem.Fields.Where(tfsField => actuallyChangedWordFields.Any(x => x.ReferenceName == tfsField.ReferenceName)).ToList());
        }

        /// <summary>
        /// Returns a list with all fields that have been changed in TFS compared to the current word revision.
        /// </summary>
        /// <param name="tfsWorkItem">Work item on TFS</param>
        /// <param name="wordWorkItem">Work item in word.</param>
        /// <param name="ignoreFormatting">Sets whether formatting is ignored.</param>
        /// <returns>List of different fields.</returns>
        public static Collection<IField> GetChangedWordFields(IWorkItem wordWorkItem, IWorkItem tfsWorkItem, bool ignoreFormatting)
        {
            Guard.ThrowOnArgumentNull(wordWorkItem, "wordWorkItem");
            Guard.ThrowOnArgumentNull(tfsWorkItem, "tfsworkItem");

            if (wordWorkItem.IsNew)
            {
                return new Collection<IField>(wordWorkItem.Fields.ToList());
            }

            var originalWordRevision = tfsWorkItem.GetWorkItemByRevision(wordWorkItem.Revision);

            return GetFieldsWithDifferentValues(wordWorkItem.Fields.ToList(), originalWordRevision.ToList(), ignoreFormatting);
        }

        /// <summary>
        /// Returns a list with all fields that have been changed in Word compared to the state saved on TFS.
        /// </summary>
        /// <param name="changeFields">List of fields that have changed in Word.</param>
        /// <param name="tfsWorkItem">Work item on TFS</param>
        /// <returns>List of different fields.</returns>
        public static Collection<IField> GetPublishableWordFields(IList<IField> changeFields, IWorkItem tfsWorkItem)
        {
            return new Collection<IField>(changeFields.Where(
                x =>
                x.Configuration.Direction != Direction.GetOnly &&
                x.Configuration.Direction != Direction.TfsToOther &&
                (x.Configuration.Direction != Direction.SetInNewTfsWorkItem || tfsWorkItem.IsNew)).ToList());
        }

        /// <summary>
        /// Returns a list with all fields that have different values for two given work items.
        /// </summary>
        /// <param name="wordFields">Fields of a Word work item.</param>
        /// <param name="tfsFields">Fields of a TFS work item.</param>
        /// <param name="ignoreFormatting">Sets whether formatting is ignored.</param>
        /// <returns>List of different fields.</returns>
        public static Collection<IField> GetFieldsWithDifferentValues2(IFieldCollection wordFields, IFieldCollection tfsFields, bool ignoreFormatting)
        {
            Guard.ThrowOnArgumentNull(wordFields, "wordFields");
            Guard.ThrowOnArgumentNull(tfsFields, "tfsFields");

            var changedFields = new Collection<IField>();

            foreach (var field in wordFields)
            {
                try
                {
                    if (tfsFields.Contains(field.ReferenceName))
                    {
                        if (!tfsFields[field.ReferenceName].CompareValue(wordFields[field.ReferenceName].Value, ignoreFormatting))
                        {
                            changedFields.Add(field);
                        }
                    }
                }
                catch (KeyNotFoundException)
                {
                    SyncServiceTrace.W("Converter value not found. Please check your configuration");
                }
            }

            return changedFields;
        }

        /// <summary>
        /// Returns a list with all fields that have different values for two given work items.
        /// </summary>
        /// <param name="newFields">Fields of a Word work item.</param>
        /// <param name="oldTfsFields">Fields of a TFS work item.</param>
        /// <param name="ignoreFormatting">Sets whether formatting is ignored.</param>
        /// <returns>List of different fields.</returns>
        public static Collection<IField> GetFieldsWithDifferentValues(IList<IField> newFields, IList<IField> oldTfsFields, bool ignoreFormatting)
        {
            Guard.ThrowOnArgumentNull(newFields, "newFields");
            Guard.ThrowOnArgumentNull(oldTfsFields, "oldTfsFields");

            var changedFields = new Collection<IField>();

            foreach (var field in newFields)
            {
                try
                {
                    var tfsField = oldTfsFields.FirstOrDefault(x => x.ReferenceName.Equals(field.ReferenceName));

                    if (tfsField != null)
                    {
                        //The right side of the if statement has been added because implementing the requirement 20904
                        if (!tfsField.CompareValue(field.Value, ignoreFormatting) || (string.IsNullOrEmpty(tfsField.Value) && field.ContainsShapes))
                        {
                            // Exclude the StackRank from the compare
                            // TODO Check again if this is necessary
                            if (!tfsField.ReferenceName.Equals("Microsoft.VSTS.Common.StackRank"))
                            {
                                changedFields.Add(field);
                            }
                        }
                    }
                }
                catch (KeyNotFoundException)
                {
                    SyncServiceTrace.W("Converter value not found. Please check your configuration");
                }
            }

            return changedFields;
        }

        /// <summary>
        /// Returns a list with all fields that have been changed in Word and TFS.
        /// </summary>
        /// <param name="changedWordFields">List with all fields that have changed in Word.</param>
        /// <param name="changedTfsFields">List with all fields that have changed in TFS.</param>
        /// <returns>List of different fields.</returns>
        public static Collection<IField> GetDivergedFields(IList<IField> changedWordFields, IList<IField> changedTfsFields)
        {
            return new Collection<IField>(changedWordFields.Where(x => changedTfsFields.Any(y => y.ReferenceName.Equals(x.ReferenceName))).ToList());
        }
    }
}
