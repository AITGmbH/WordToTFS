#region Usings
using System;
using System.Collections;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.WorkItemCollections;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Contracts.Exceptions;
using System.Globalization;
using AIT.TFS.SyncService.Factory;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemCollections
{
    /// <summary>
    /// Class implements the collection interface <see cref="IFieldCollection"/>.
    /// </summary>
    public class TfsFieldCollection : IFieldCollection
    {
        private readonly IDictionary<string, IField> _internalCollection = new Dictionary<string, IField>();

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsFieldCollection"/> class.
        /// </summary>
        /// <param name="fields">Fields to add to the collection.</param>
        internal TfsFieldCollection(IEnumerable<IField> fields)
        {
            foreach (IField field in fields)
                _internalCollection.Add(field.ReferenceName.ToUpperInvariant(), field);
        }

        #region IFieldCollection Members

        /// <summary>
        /// Access a specific field via refName
        /// </summary>
        /// <param name="refName">The reference name of the field</param>
        /// <returns>The referenced field</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Justification = "Not an unexpected location. Replaces a less informative exception.")]
        public IField this[string refName]
        {
            get
            {
                if (!string.IsNullOrEmpty(refName) && _internalCollection.ContainsKey(refName.ToUpperInvariant()))
                {
                    return _internalCollection[refName.ToUpperInvariant()];
                }

                throw new ConfigurationException($"The Field {refName} does not exist or is not mapped for your work item type in the WordToTFS template. Check your Field definition and spelling.");
            }
        }

        /// <summary>
        /// Determines whether are give refName exists in the collection
        /// </summary>
        /// <param name="refName">The reference name</param>
        /// <returns>true if it exists / false otherwise</returns>
        public bool Contains(string refName)
        {
            Guard.ThrowOnArgumentNull(refName, "refName");
            return _internalCollection.ContainsKey(refName.ToUpperInvariant());
        }

        /// <summary>
        /// Adds an item to the ICollection&lt;T&gt;.
        /// </summary>
        /// <param name="item">The object to add to the ICollection&lt;T&gt;.</param>
        public void Add(IField item)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Removes all items from the ICollection&lt;T&gt;.
        /// </summary>
        public void Clear()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Determines whether the System.Collections.Generic.ICollection&lt;T&gt; contains a specific value.
        /// </summary>
        /// <param name="item">The object to locate in the ICollection&lt;T&gt;.</param>
        /// <returns>true if item is found in the ICollection&lt;T&gt;; otherwise, false.</returns>
        public bool Contains(IField item)
        {
            if (item == null) throw new ArgumentNullException("item");
            return Contains(item.ReferenceName);
        }

        /// <summary>
        /// Gets the number of elements contained in the ICollection&lt;T&gt;.
        /// </summary>
        public int Count
        {
            get { return _internalCollection.Count; }
        }

        /// <summary>
        /// Copies the elements of the ICollection&lt;T&gt; to an Array, starting at a particular Array index.
        /// </summary>
        /// <param name="array">The one-dimensional Array that is the destination of the elements copied from ICollection&lt;T&gt;. The Array must have zero-based indexing.</param>
        /// <param name="arrayIndex">The zero-based index in array at which copying begins.</param>
        public void CopyTo(IField[] array, int arrayIndex)
        {
            _internalCollection.Values.CopyTo(array, arrayIndex);
        }

        /// <summary>
        /// Gets a value indicating whether the ICollection&lt;T&gt; is read-only.
        /// </summary>
        public bool IsReadOnly
        {
            get { return false; }
        }

        /// <summary>
        /// Removes the first occurrence of a specific object from the ICollection&lt;T&gt;.
        /// </summary>
        /// <param name="item">The object to remove from the ICollection&lt;T&gt;.</param>
        /// <returns></returns>
        public bool Remove(IField item)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>An IEnumerator object that can be used to iterate through the collection.</returns>
        public IEnumerator GetEnumerator()
        {
            return _internalCollection.Values.GetEnumerator();
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>A IEnumerator&lt;T&gt; that can be used to iterate through the collection.</returns>
        IEnumerator<IField> IEnumerable<IField>.GetEnumerator()
        {
            return _internalCollection.Values.GetEnumerator();
        }

        #endregion
    }
}