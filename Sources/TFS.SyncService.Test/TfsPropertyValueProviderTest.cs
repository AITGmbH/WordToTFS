#region Usings
using System;
using System.Collections.Generic;
using System.Globalization;
using AIT.TFS.SyncService.Adapter.TFS2012.TestCenter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
#endregion

namespace TFS.SyncService.Test.Unit
{
    /// <summary>
    /// Tests the PropertyValueProvider class
    /// </summary>
    [TestClass]
    public class TfsPropertyValueProviderTest
    {
        #region TestMethods
        /// <summary>
        /// Test bindings like <code>ExecutionDate.ToString("dd-MMM-YYYY")</code>
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_CallMethodWithStringParameter_ShouldEvaluateCorrectly()
        {
            // setup
            var tfsPropertyValueProvider = new Mock<TfsPropertyValueProvider>
            {
                CallBase = true
            };

            tfsPropertyValueProvider.SetupGet(x => x.AssociatedObject).Returns(new TestClass());

            // test
            var result = tfsPropertyValueProvider.Object.PropertyValue("MethodWithStringParameter(\"ActuallyPassedParameter\")", x => x);

            // verify
            Assert.AreEqual("ActuallyPassedParameter", result);
        }

        /// <summary>
        /// Test bindings like <code>ExecutionDate.ToString("dd-MMM-YYYY")</code>
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_CallMethodWithIntParameter_ShouldEvaluateCorrectly()
        {
            // setup
            var tfsPropertyValueProvider = new Mock<TfsPropertyValueProvider>
            {
                CallBase = true
            };

            tfsPropertyValueProvider.SetupGet(x => x.AssociatedObject).Returns(new TestClass());

            // test
            var result = tfsPropertyValueProvider.Object.PropertyValue("MethodWithIntParameter(4)", x => x);

            // verify
            Assert.AreEqual("4", result);
        }

        /// <summary>
        /// Test calling a method with comma separated parameters
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_CallMethodWithMultipleParameters_ShouldEvaluateCorrectly()
        {
            // setup
            var parameter1 = "5";
            var parameter2 = "test";
            var parameter3 = "3.14";
            var tfsPropertyValueProvider = new Mock<TfsPropertyValueProvider>
            {
                CallBase = true
            };

            tfsPropertyValueProvider.SetupGet(x => x.AssociatedObject).Returns(new TestClass());

            // test
            var parameters = string.Join(",", parameter1, parameter2, parameter3);
            var result = tfsPropertyValueProvider.Object.PropertyValue("MethodWithMultipleParameters(" + parameters + ")", x => x);

            // verify
            var expected = string.Join("#", parameter1, parameter2, parameter3);
            Assert.AreEqual(expected, result);
        }

        /// <summary>
        /// Test bindings like ExecutionDate.ToString()
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_CallMethodWithoutParameter_ShouldEvaluateCorrectly()
        {
            // setup
            var tfsPropertyValueProvider = new Mock<TfsPropertyValueProvider>
            {
                CallBase = true
            };

            tfsPropertyValueProvider.SetupGet(x => x.AssociatedObject).Returns(new TestClass());

            // test
            var result = tfsPropertyValueProvider.Object.PropertyValue("MethodWithoutParameter()", x => x);

            // verify
            Assert.AreEqual("MethodWithoutParameterResult", result);
        }

        /// <summary>
        /// Test bindings like ExecutionDate.ToString()
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_GetProperty_ShouldEvaluateCorrectly()
        {
            // setup
            var tfsPropertyValueProvider = new Mock<TfsPropertyValueProvider>
            {
                CallBase = true
            };

            tfsPropertyValueProvider.SetupGet(x => x.AssociatedObject).Returns(new TestClass());

            // test
            var result = tfsPropertyValueProvider.Object.PropertyValue("MyTestProperty", x => x);

            // verify
            Assert.AreEqual("PropertyValue", result);
        }

        /// <summary>
        /// Test dictionary access
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_GetStringIndexedProperty_ShouldEvaluateCorrectly()
        {
            // setup
            var tfsPropertyValueProvider = new Mock<TfsPropertyValueProvider>
            {
                CallBase = true
            };

            tfsPropertyValueProvider.SetupGet(x => x.AssociatedObject).Returns(new TestClass());

            // test
            var result = tfsPropertyValueProvider.Object.PropertyValue("DictionaryProperty[\"OnlyKey\"]", x => x);

            // verify
            Assert.AreEqual("OnlyValue", result);
        }

        /// <summary>
        /// Test dictionary access
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_GetIntIndexedProperty_ShouldEvaluateCorrectly()
        {
            // setup
            var tfsPropertyValueProvider = new Mock<TfsPropertyValueProvider>
            {
                CallBase = true
            };

            tfsPropertyValueProvider.SetupGet(x => x.AssociatedObject).Returns(new TestClass());

            // test
            var result = tfsPropertyValueProvider.Object.PropertyValue("ListProperty[1]", x => x);

            // verify
            Assert.AreEqual("SecondListEntry", result);
        }

        /// <summary>
        /// Test array access
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_GetArrayProperty_ShouldEvaluateCorrectly()
        {
            // setup
            var tfsPropertyValueProvider = new Mock<TfsPropertyValueProvider>
            {
                CallBase = true
            };

            tfsPropertyValueProvider.SetupGet(x => x.AssociatedObject).Returns(new TestClass());

            // test
            var result = tfsPropertyValueProvider.Object.PropertyValue("ArrayProperty[1]", x => x);

            // verify
            Assert.AreEqual("SecondArrayPropertyValue", result);
        }

        /// <summary>
        /// Test if the wrapper object correctly hides properties of the replaced object.
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_AccessShadowedProperty_ShouldReturnWrapperValue()
        {
            // setup
            var tfsPropertyValueProvider = new TestClassWrapper(new TestClass());

            // test
            var result = tfsPropertyValueProvider.PropertyValue("MyTestProperty", x => x);

            // verify
            Assert.AreEqual("WrapperPropertyValue", result);
        }

        /// <summary>
        /// Test if the wrapper object correctly hides properties of the replaced object.
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_AccessShadowedMethod_ShouldReturnWrapperValue()
        {
            // setup
            var tfsPropertyValueProvider = new TestClassWrapper(new TestClass());

            // test
            var result = tfsPropertyValueProvider.PropertyValue("MethodWithStringParameter(\"ActuallyPassedParameter\")", x => x);

            // verify
            Assert.AreEqual("WrapperMethodWithStringParameterResult", result);
        }

        /// <summary>
        /// Test if the wrapper object correctly hides properties of the replaced object.
        /// </summary>
        [TestMethod]
        public void TestCenter_TfsPropertyValueProvider_CallMethodsTest_ShouldReturnFormattedDate()
        {
            // setup
            var tfsPropertyValueProvider = new TestClassWrapper(new TestClass());

            // test
            var result = tfsPropertyValueProvider.PropertyValue("Date.ToString(\"dd-MMM-yyyy\")", x => x);

            // verify
            Assert.AreEqual("05-Jan-1999", result);
        }
        #endregion

        #region Test Classes
        private class TestClassWrapper : TfsPropertyValueProvider
        {
            private readonly TestClass _testClass;

            public override object AssociatedObject
            {
                get
                {
                    return _testClass;
                }
            }

            public TestClassWrapper(TestClass testClass)
            {
                _testClass = testClass;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public string MyTestProperty
            {
                get
                {
                    return "WrapperPropertyValue";
                }
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "parameter", Justification = "Parameter value not important")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public string MethodWithStringParameter(string parameter)
            {
                return "WrapperMethodWithStringParameterResult";
            }
        }

        private class TestClass
        {
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public DateTime Date
            {
                get
                {
                   return new DateTime(1999, 1, 5);
                }
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public Dictionary<string, string> DictionaryProperty
            {
                get
                {
                    return new Dictionary<string, string>
                               {
                                   {
                                       "OnlyKey", "OnlyValue"
                                   }
                               };
                }
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public List<string> ListProperty
            {
                get
                {
                    return new List<string>
                               {
                                   "FirstListEntry",
                                   "SecondListEntry"
                               };
                }
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public string[] ArrayProperty
            {
                get
                {
                    return new []
                               {
                                   "FirstArrayPropertyValue", "SecondArrayPropertyValue"
                               };
                }
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public string MyTestProperty
            {
                get
                {
                    return "PropertyValue";
                }
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public string MethodWithStringParameter(string parameter)
            {
                return parameter;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public int MethodWithIntParameter(int parameter)
            {
                return parameter;
            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public string MethodWithMultipleParameters(int parameter1, string parameter2, double parameter3)
            {
                return string.Join("#", parameter1.ToString(CultureInfo.InvariantCulture), parameter2, parameter3.ToString(CultureInfo.InvariantCulture));

            }

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Need an instance member for testing.")]
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Used via reflection.")]
            public string MethodWithoutParameter()
            {
                return "MethodWithoutParameterResult";
            }

        }

        #endregion
    }
}
