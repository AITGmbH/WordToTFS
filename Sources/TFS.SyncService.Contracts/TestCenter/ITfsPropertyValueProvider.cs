using System;
using System.Collections;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    /// <summary>
    /// Interface defines functionality of team foundation server test object.
    /// </summary>
    public interface ITfsPropertyValueProvider
    {
        /// <summary>
        /// The method determines value of specified property in original object of test case.
        /// </summary>
        /// <param name="propertyName">Name of property whose value is to be determined.</param>
        /// <param name="enumerableExpander">The method expands enumerable.</param>
        /// <returns>Determined value of required property.</returns>
        string PropertyValue(string propertyName, Func<IEnumerable, IEnumerable> enumerableExpander);

        /// <summary>
        /// The method determines value of specified property in original object of test case.
        /// </summary>
        /// <param name="propertyName">Name of property whose value is to be determined.</param>
        /// <param name="valueConsumer">Consumer of provided property value.</param>
        /// <param name="enumerableExpander">The method expands enumerable.</param>
        void PropertyValue(string propertyName, Action<object> valueConsumer, Func<IEnumerable, IEnumerable> enumerableExpander);

        /// <summary>
        /// The method creates one property value provider with given object.
        /// </summary>
        /// <param name="propertyProvider">Object contains the examined properties.</param>
        /// <returns>Temporary value provider.</returns>
        ITfsPropertyValueProvider GetTemporaryPropertyValueProvider(object propertyProvider);

        /// <summary>
        /// Gets the object that is wrapped by the value provider
        /// </summary>
        object AssociatedObject { get; }

        /// <summary>
        /// Gets a custom object. This is used to retrived the values set by object queries
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        object GetCustomObject(string name);

        /// <summary>
        /// Add a custom object.This is used to store the values of the object queries
        /// </summary>
        /// <param name="name">The name of the object</param>
        /// <param name="propertyValue">The stored object</param>
        void AddCustomObjects(string name, object propertyValue);
    }
}
