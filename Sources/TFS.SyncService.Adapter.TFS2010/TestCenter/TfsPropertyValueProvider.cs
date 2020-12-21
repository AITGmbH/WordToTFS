#region Usings
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using AIT.TFS.SyncService.Adapter.TFS2012.Properties;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.TestCenter;
using AIT.TFS.SyncService.Factory;
using Microsoft.TeamFoundation.TestManagement.Client;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// Base class for all detail classes in TestCenter namespace.
    /// </summary>
    public abstract class TfsPropertyValueProvider : ITfsPropertyValueProvider
    {
        #region Consts
        private const char ConstPropertyDelimiterChar = '.';
        private const char ConstQuetesChar = '\"';
        private const char ConstLeftArrayBracketChar = '[';
        private const char ConstRightArrayBracketChar = ']';
        private const char ConstLeftBraceChar = '(';
        private const char ConstRightBraceChar = ')';
        private const string ConstPropertyDelimiterString = ".";
        private const string ConstValueDelimiter = ", ";
        #endregion

        #region Fields
        //Store any customized objects this property allwos the storage of WorkItems for TestResults and makes them accessilbe by the reporting
        private Dictionary<string, object> _customObjects = new Dictionary<string, object>();
        #endregion

        #region Properties

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public abstract object AssociatedObject { get; }
        #endregion

        #region Public methods

        /// <summary>
        /// Get a custom object from the dictionary
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public object GetCustomObject(string name)
        {
            if (_customObjects.ContainsKey(name))
            {
                return _customObjects[name];

            }
            return null;
        }

        /// <summary>
        /// Add a custom objects to the dictionary.
        /// </summary>
        /// <param name="name">The name of the stored object</param>
        /// <param name="propertyValue">The stored object</param>
        public void AddCustomObjects(string name, object propertyValue)
        {
            if (!_customObjects.ContainsKey(name))
                _customObjects.Add(name, propertyValue);
        }


        /// <summary>
        /// The method determines value of the property in <see cref="AssociatedObject"/> object.
        /// </summary>
        /// <param name="propertyName">Name of the property to determine.</param>
        /// <param name="enumerableExpander">The method expands enumerable.</param>
        /// <returns>Determined value of property in associated object.</returns>
        public string PropertyValue(string propertyName, Func<IEnumerable, IEnumerable> enumerableExpander)
        {
            Guard.ThrowOnArgumentNull(enumerableExpander, "enumerableExpander");

            if (string.IsNullOrEmpty(propertyName))
                return string.Empty;

            try
            {
                return PropertyValue(this, propertyName, enumerableExpander);
            }
            catch (Exception ex)
            {
                throw new ConfigurationException(string.Format(CultureInfo.CurrentCulture, Resources.TestReport_DetermineValueFailed, propertyName, AssociatedObject.GetType().Name), ex);
            }
        }

        /// <summary>
        /// The method determines value of specified property in original object of test case.
        /// </summary>
        /// <param name="propertyName">Name of property whose value is to be determined.</param>
        /// <param name="valueConsumer">Consumer of provided property value.</param>
        /// <param name="enumerableExpander">The method expands enumerable.</param>
        public void PropertyValue(string propertyName, Action<object> valueConsumer, Func<IEnumerable, IEnumerable> enumerableExpander)
        {
            Guard.ThrowOnArgumentNull(valueConsumer, "valueConsumer");
            Guard.ThrowOnArgumentNull(enumerableExpander, "enumerableExpander");

            if (string.IsNullOrEmpty(propertyName)) return;

            try
            {
                PropertyValue(this, propertyName, valueConsumer, enumerableExpander);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new ConfigurationException(string.Format(CultureInfo.CurrentCulture, Resources.TestReport_DetermineValueFailed, propertyName, AssociatedObject.GetType().Name), ex);
            }
        }

        /// <summary>
        /// The method creates one property value provider with given object.
        /// </summary>
        /// <param name="propertyProvider">Object contains the examined properties.</param>
        /// <returns>Temporary value provider.</returns>
        public ITfsPropertyValueProvider GetTemporaryPropertyValueProvider(object propertyProvider)
        {
            Guard.ThrowOnArgumentNull(propertyProvider, "propertyProvider");

            var temporaryPropertyValueProvider = propertyProvider as TfsPropertyValueProvider;
            return temporaryPropertyValueProvider ?? new TfsTemporaryProvider(propertyProvider);
        }
        #endregion

        #region Private methods

        /// <summary>
        /// The method determines value of given property in given object
        /// </summary>
        /// <param name="propertyContainer">Object containing the property.</param>
        /// <param name="propertyName">Property to determine value for.</param>
        /// <param name="valueConsumer">Consumer of provided property value.</param>
        /// <param name="enumerableExpander">The method expands enumerable.</param>
        /// <returns>Determined value.</returns>
        private void PropertyValue(object propertyContainer, string propertyName, Action<object> valueConsumer, Func<IEnumerable, IEnumerable> enumerableExpander)
        {
            Guard.ThrowOnArgumentNull(propertyName, "propertyName");
            Guard.ThrowOnArgumentNull(valueConsumer, "valueConsumer");

            var parts = SplitPathIntoProperties(propertyName);
            var value = GetValueFromPropertyContainer(propertyContainer, parts[0]);

            if (parts.Length == 1)
            {
                // 'null' is legal value
                if (value == null)
                {
                    valueConsumer(null);
                    return;
                }
                // Check IEnumerable property
                if (IsEnumerable(value))
                {
                    // Handle enumerable property
                    // Don't check "value as IEnumerable".
                    // User defines this property as IEnumerable.
                    foreach (var oneValue in enumerableExpander(value as IEnumerable))
                    {
                        valueConsumer(oneValue);
                    }
                    return;
                }
                // Handle scalar property
                valueConsumer(value);
            }
            else
            {
                var path = parts.Skip(1).Aggregate((x, y) => x + ConstPropertyDelimiterString + y);
                PropertyValue(value, path, valueConsumer, enumerableExpander);
            }
        }

        /// <summary>
        /// The method determines value of given property in given object
        /// </summary>
        /// <param name="propertyContainer">Object containing the property.</param>
        /// <param name="propertyName">Property to determine value for.</param>
        /// <param name="enumerableExpander">The method expands enumerable.</param>
        /// <returns>Determined value.</returns>
        private string PropertyValue(object propertyContainer, string propertyName, Func<IEnumerable, IEnumerable> enumerableExpander)
        {
            Guard.ThrowOnArgumentNull(propertyName, "propertyName");

            var parts = SplitPathIntoProperties(propertyName);
            var value = GetValueFromPropertyContainer(propertyContainer, parts[0]);


            if (parts.Length == 1)
            {
                // 'null' is legal value
                if (value == null) return string.Empty;

                // If the value is a parameterized string (which is also an IEnumerable, 
                // but it enumerates parts that have no reasonable ToString method), 
                // then just return the string without parameters.
                var parameterizedString = value as ParameterizedString;
                if (parameterizedString != null)
                {
                    return parameterizedString.ToString();
                }

                // Check IEnumerable property
                if (IsEnumerable(value))
                {
                    // Handle enumerable property
                    var allElements = string.Empty;
                    // Don't check "value as IEnumerable".
                    // User defines this property as IEnumerable.
                    foreach (var oneValue in enumerableExpander(value as IEnumerable))
                    {
                        if (allElements.Length > 0)
                        {
                            allElements += ConstValueDelimiter;
                        }
                        allElements += oneValue;
                    }
                    return allElements;
                }
                // Handle scalar value
                return value.ToString();
            }


            return PropertyValue(value, parts.Skip(1).Aggregate((x, y) => x + ConstPropertyDelimiterString + y), enumerableExpander);
        }

        /// <summary>
        /// Splits a binding path like System.Console.BufferHeight into the tree components System, Console and BufferHeight
        /// </summary>
        private static string[] SplitPathIntoProperties(string path)
        {
            var components = path.Split(ConstPropertyDelimiterChar);
            var mergedComponents = new List<string>();
            var openQuotationMarks = false;
            var openMethodBraces = false;

            // string split splits indexed properties that contain a perdiod like "Fields["System.Description"]"
            // although its actually only one property.
            for (int i = 0; i < components.Length; i++)
            {
                if (openQuotationMarks || openMethodBraces)
                {
                    var last = mergedComponents.Last();
                    mergedComponents.Remove(last);
                    mergedComponents.Add(last + ConstPropertyDelimiterChar + components[i]);
                }
                else
                {
                    mergedComponents.Add(components[i]);
                }

                openMethodBraces = components[i].Count(x => x == '(') > components[i].Count(x => x == ')');
                if (components[i].Count(x => x == '"') % 2 == 1)
                {
                    openQuotationMarks = !openQuotationMarks;
                }
            }

            if (openMethodBraces) throw new ConfigurationException($"The path '{path}' is malformed. Make sure all opening braces are closed.");
            if (openQuotationMarks) throw new ConfigurationException($"The path '{path}' is malformed. Make sure all quotation marks are properly set.");

            return mergedComponents.ToArray();
        }

        /// <summary>
        /// The method evaluates property / method in given object.
        /// </summary>
        /// <param name="propertyContainer">Object to evaluate the property / method.</param>
        /// <param name="propertyName">Name of property / method.</param>
        /// <returns>Returned value from property / method.</returns>
        private static object GetValueFromPropertyContainer(object propertyContainer, string propertyName)
        {
            var valueProvider = propertyContainer as TfsPropertyValueProvider;

            var arrayPropertyName = string.Empty;
            int? elementIndex;
            var elementName = string.Empty;
            if (ParseArrayProperty(propertyName, out arrayPropertyName, out elementIndex, out elementName))
            {
                var arrayPropertyInfo = propertyContainer.GetType().GetProperty(arrayPropertyName);
                if (arrayPropertyInfo == null)
                {
                    if (valueProvider != null)
                    {
                        arrayPropertyInfo = valueProvider.AssociatedObject.GetType().GetProperty(arrayPropertyName);
                        if (arrayPropertyInfo != null)
                        {
                            return GetValueFromIndexedProperty(valueProvider.AssociatedObject, arrayPropertyInfo, elementIndex, elementName);
                        }
                    }
                }
                else
                {
                    return GetValueFromIndexedProperty(propertyContainer, arrayPropertyInfo, elementIndex, elementName);
                }

                throw new ConfigurationException(string.Format(CultureInfo.InvariantCulture, "The object of type '{0}' does not contain an indexed property named '{1}'", (valueProvider == null ? propertyContainer : valueProvider.AssociatedObject).GetType(), propertyName));
            }

            var methodName = string.Empty;
            object[] parameter;
            if (ParseMethodCall(propertyName, out methodName, out parameter))
            {
                var methodInfo = propertyContainer.GetType().GetMethod(methodName, parameter.Select(x => x.GetType()).ToArray());
                if (methodInfo == null)
                {
                    if (valueProvider != null)
                    {
                        methodInfo = valueProvider.AssociatedObject.GetType().GetMethod(methodName, parameter.Select(x => x.GetType()).ToArray());
                        if (methodInfo != null)
                        {
                            return methodInfo.Invoke(valueProvider.AssociatedObject, parameter);
                        }
                    }
                }
                else
                {
                    return methodInfo.Invoke(propertyContainer, parameter);
                }
            }

            // The property is something like 'Name' or 'Year'. Try to find on wrapper object first
            var propertyInfo = propertyContainer.GetType().GetProperty(propertyName);
            if (propertyInfo != null)
                return propertyInfo.GetValue(propertyContainer, null);

            if (valueProvider != null)
            {
                //Check the custom properties 
                if (valueProvider.GetCustomObject(propertyName) != null)
                {
                    return valueProvider.GetCustomObject(propertyName);
                }

                // The property is something like 'Name' or 'Year'. Try to find on wrapped object
                propertyInfo = valueProvider.AssociatedObject.GetType().GetProperty(propertyName);
                if (propertyInfo != null)
                {
                    return propertyInfo.GetValue(valueProvider.AssociatedObject, null);
                }
            }

            //TODO MIS 2015-07-16: Find better ways. Currently, if property cannot be found by using the current object type which might be based on TFS Report API, ID will be used to query the TFS WI API
            //TODO MIS 2015-07-16: Central static solution for querying WIs via Store by Id etc.
            if (valueProvider != null && (valueProvider.AssociatedObject is ITestObject<int>))
            {
                var testCase = (ITestObject<int>)valueProvider.AssociatedObject;
                var workItemStore = testCase.Project.WitProject.Store;
                var propertyIdInfo = valueProvider.AssociatedObject.GetType().GetProperty("Id");
                var currentWiId = (int)propertyIdInfo.GetValue(valueProvider.AssociatedObject, null);
                var workItem = workItemStore.GetWorkItem(currentWiId);
                if (workItem.Fields.Contains(propertyName))
                {
                    var value = workItem.Fields[propertyName].Value;
                    return value;
                }
            }

            if (valueProvider != null && valueProvider.AssociatedObject is ITestAttachment)
            {
                return (valueProvider.AssociatedObject as ITestAttachment).Name;
            }
            
            throw new ConfigurationException(
                string.Format(CultureInfo.InvariantCulture, "The object of type '{0}' does not contain a property named '{2}'"
                , (valueProvider == null ? propertyContainer : valueProvider.AssociatedObject).GetType()
                , propertyName));
        }

        private static object GetValueFromIndexedProperty(object propertyContainer, PropertyInfo propertyInfo, int? elementIndex, string elementName)
        {
            // The property is something like 'Names[3]' or 'Field["Description"]'
            // Create parameters for property
            var parameters = new List<object>();
            if (elementIndex != null)
            {
                parameters.Add(elementIndex.Value);
            }
            else if (!string.IsNullOrEmpty(elementName))
            {
                parameters.Add(elementName);
            }

            // ... and parameters of the property.
            var pars = propertyInfo.GetIndexParameters();
            if (pars.Length == 1)
            {
                // The property can be direct indexed.
                try
                {
                    return propertyInfo.GetValue(propertyContainer, parameters.ToArray());
                }
                catch
                {
                    throw new ConfigurationException($"The key or index '{parameters.First()}' was not found in the collection '{propertyInfo.Name}'");
                }
            }
            // We need to get the value of property and then use the index.
            var arrayObject = propertyInfo.GetValue(propertyContainer, null);
            PropertyInfo indexProperty = null;
            if (elementIndex != null)
            {
                indexProperty = arrayObject.GetType().GetProperty("Item", new[] { typeof(int) });
            }
            else if (!string.IsNullOrEmpty(elementName))
            {
                indexProperty = arrayObject.GetType().GetProperty("Item", new[] { typeof(string) });
            }
            if (indexProperty != null)
            {
                try
                {
                    return indexProperty.GetValue(arrayObject, parameters.ToArray());
                }
                catch
                {
                    throw new ConfigurationException($"The key or index '{parameters.First()}' was not found in the collection '{propertyInfo.Name}'");
                }
            }

            if (arrayObject is Array)
            {
                try
                {
                    return (arrayObject as Array).GetValue(parameters.Cast<int>().ToArray());
                }
                catch
                {
                    throw new ConfigurationException($"The key or index '{parameters.First()}' was not found in the collection '{propertyInfo.Name}'");
                }
            }

            throw new ConfigurationException($"Cannot access property '{propertyInfo.Name}' using index or key.");
        }

        /// <summary>
        /// The method determines if an object is enumerable
        /// </summary>
        /// <param name="value"><see cref="PropertyInfo"/> to check.</param>
        private static bool IsEnumerable(object value)
        {
            // We don't consider a string enumerable as there is no useful per-character presentation
            if (value is string) return false;
            if (value.GetType().GetInterfaces().Any(t => t.IsGenericType && t.GetGenericTypeDefinition() == typeof(IEnumerable<>))) return true;
            if (value.GetType().GetInterfaces().Any(t => t == typeof(IEnumerable))) return true;

            return false;
        }

        /// <summary>
        /// The method determines whether the property is array property.
        /// For example 'Names[3]' or 'Field["Description"]'
        /// </summary>
        /// <param name="property">The property to examine.</param>
        /// <param name="arrayPropertyName">Returns not <c>null</c> if the property is array property.
        /// Here is the right name of property without brackets.</param>
        /// <param name="elementIndex">Returns not <c>null</c> if the array property has numeric index.</param>
        /// <param name="elementName">Returns not <c>null</c> if the array property has text index.</param>
        /// <returns><c>true</c> if the property is array property
        /// and <paramref name="elementIndex"/>, or <paramref name="elementName"/> is set.
        /// Otherwise <c>false</c> and both parameters are set to <c>null</c>.</returns>
        private static bool ParseArrayProperty(string property, out string arrayPropertyName, out int? elementIndex, out string elementName)
        {
            elementIndex = null;
            elementName = null;
            arrayPropertyName = null;
            if (string.IsNullOrEmpty(property))
                return false;
            var leftPosition = property.IndexOf(ConstLeftArrayBracketChar);
            var rightPosition = property.IndexOf(ConstRightArrayBracketChar);
            if (leftPosition < 0 || rightPosition < 0 || leftPosition > rightPosition)
                return false;
            arrayPropertyName = property.Substring(0, leftPosition);
            var indexText = property.Substring(leftPosition + 1, rightPosition - leftPosition - 1);
            var tryInt = 0;
            if (int.TryParse(indexText, out tryInt))
            {
                elementIndex = tryInt;
                return true;
            }
            elementName = indexText.Trim(ConstQuetesChar);
            return true;
        }


        private static bool ParseMethodCall(string property, out string methodName, out object[] parameter)
        {
            methodName = null;
            parameter = null;

            if (string.IsNullOrEmpty(property))
            {
                return false;
            }

            var leftPosition = property.IndexOf(ConstLeftBraceChar);
            var rightPosition = property.IndexOf(ConstRightBraceChar);
            if (leftPosition < 0 || rightPosition < 0 || leftPosition > rightPosition)
            {
                return false;
            }

            methodName = property.Substring(0, leftPosition);
            var indexText = property.Substring(leftPosition + 1, rightPosition - leftPosition - 1);
            var stringParameters = indexText.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim().Replace("\"", string.Empty)).ToArray();
            parameter = new object[stringParameters.Length];
            for (int i = 0; i < stringParameters.Length; i++)
            {
                var intValue = 0;
                double doubleValue;

                if (Int32.TryParse(stringParameters[i], out intValue))
                {
                    parameter[i] = intValue;
                }
                else if (Double.TryParse(stringParameters[i], NumberStyles.Any, CultureInfo.InvariantCulture, out doubleValue))
                {
                    parameter[i] = doubleValue;
                }
                else
                {
                    parameter[i] = stringParameters[i];
                }
            }

            return true;
        }
        #endregion


    }
}
