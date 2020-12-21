using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;

namespace AIT.TFS.SyncService.Contracts.WorkItemCollections
{
    /// <summary>
    /// Interface defines functionality of <see cref="IField"/> collection.
    /// </summary>
    public interface IFieldCollection : ICollection<IField>
    {
        /// <summary>
        /// Access a specific field via refName
        /// </summary>
        /// <param name="refName">the reference name of the field</param>
        /// <returns>The referenced field</returns>
        IField this[string refName] { get; }

        /// <summary>
        /// Determines whether are give refName exists in the collection
        /// </summary>
        /// <param name="refName">The reference name</param>
        /// <returns>True if it exists / false otherwise</returns>
        bool Contains(string refName);
    }
}