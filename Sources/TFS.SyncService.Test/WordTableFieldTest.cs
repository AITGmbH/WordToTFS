#region Usings
using System.Collections.Generic;
using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
#endregion

namespace TFS.SyncService.Test.Unit
{

    /// <summary>
    /// Integration test for the WordTableField adapter and its word document.
    /// </summary>
    [TestClass]
    public class WordTableFieldTest
    {
        #region Fields
        private Document _testDocument;
        private Microsoft.Office.Interop.Word.Range _cellRange;
        private Mock<IConfigurationFieldItem> _fieldConfiguration;
        private Mock<IConverter> _converter;
        #endregion

        #region Test initializations

        /// <summary>
        /// Create document and mock configuration
        /// </summary>
        [TestInitialize]
        public void TestInitialize()
        {
            TFS.Test.Common.TestCleanup.CloseWordDocumentAndKillOpenWordInstances();
            _testDocument = new Document();
            _testDocument.Tables.Add(_testDocument.ActiveWindow.Document.Range(), 1, 1);
            _cellRange = _testDocument.Tables[1].Cell(1, 1).Range;
            _fieldConfiguration = new Mock<IConfigurationFieldItem>();
            _fieldConfiguration.Setup(x => x.Clone()).Returns(_fieldConfiguration.Object);
            _converter = new Mock<IConverter>();
        }

        /// <summary>
        /// Close document
        /// </summary>
        [TestCleanup]
        public void TestCleanup()
        {
            if (_testDocument != null)
            {
                TFS.Test.Common.TestCleanup.CloseWordDocument(_testDocument);
            }
        }
        #endregion

        #region TestMethods

        /// <summary>
        /// Should create DropDown list.
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableField_SetAllowedValues_ShouldCreateDropDownList()
        {
            _fieldConfiguration.SetupGet(x => x.FieldValueType).Returns(FieldValueType.DropDownList);

            var wordField = new WordTableField(_cellRange, _fieldConfiguration.Object, _converter.Object, true);
            wordField.AllowedValues = new List<string>
                                          {
                                              "A",
                                              "B",
                                              "C"
                                          };

            Assert.IsTrue(_cellRange.ContentControls[1].Type == WdContentControlType.wdContentControlDropdownList);
            Assert.AreEqual("A", _cellRange.ContentControls[1].DropdownListEntries[1].Text);
            Assert.AreEqual("C", _cellRange.ContentControls[1].DropdownListEntries[3].Text);
        }

        /// <summary>
        /// Don't add DropDown list if one exists
        /// </summary>
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableField_SetAllowedValues_ShouldNotAddDropDownListIfOneExists()
        {
            _cellRange.ContentControls.Add(WdContentControlType.wdContentControlDropdownList);
            _fieldConfiguration.SetupGet(x => x.FieldValueType).Returns(FieldValueType.DropDownList);
            var wordField = new WordTableField(_cellRange, _fieldConfiguration.Object, _converter.Object, true);
            wordField.AllowedValues = new List<string>
                                          {
                                              "A",
                                              "B",
                                              "C"
                                          };

            Assert.AreEqual(1, _cellRange.ContentControls.Count);
        }

        /// <summary>
        /// Check if setting the value in the wrapper actually changes the text in the cell
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "Microsoft.Office.Interop.Word.ContentControlListEntries.Add(System.String,System.String,System.Int32)")]
        [TestMethod]
        [TestCategory("Interactive")]
        public void WordTableField_SetValue_ShouldSelectItemInDropDownList()
        {
            _cellRange.ContentControls.Add(WdContentControlType.wdContentControlDropdownList);
            _cellRange.ContentControls[1].DropdownListEntries.Add("A", "A");
            _cellRange.ContentControls[1].DropdownListEntries.Add("B", "B");

            _fieldConfiguration.SetupGet(x => x.FieldValueType).Returns(FieldValueType.DropDownList);

            var wordField = new WordTableField(_cellRange, _fieldConfiguration.Object, _converter.Object, true);
            wordField.Value = "B";

            Assert.IsTrue(_cellRange.ContentControls[1].Type == WdContentControlType.wdContentControlDropdownList);
            Assert.AreEqual("B", _cellRange.Text);
        }

        /// <summary>
        /// Check if setting the value of a dropdown list to a value not in the list does nothing
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "Microsoft.Office.Interop.Word.ContentControlListEntries.Add(System.String,System.String,System.Int32)")]
        public void WordTableField_SetValue_ShouldDoNothingIfItemNotInDropDownList()
        {
            _cellRange.ContentControls.Add(WdContentControlType.wdContentControlDropdownList);
            _cellRange.ContentControls[1].DropdownListEntries.Add("A", "A");
            _cellRange.ContentControls[1].DropdownListEntries[1].Select();

            _fieldConfiguration.SetupGet(x => x.FieldValueType).Returns(FieldValueType.DropDownList);

            var wordField = new WordTableField(_cellRange, _fieldConfiguration.Object, _converter.Object, true);
            wordField.Value = "B";

            Assert.AreEqual("A", _cellRange.Text);
        }
        #endregion
    }
}
