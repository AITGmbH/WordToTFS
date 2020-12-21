using System.Collections.Generic;
using System.Text.RegularExpressions;
using AIT.TFS.SyncService.Adapter.Word2007.Properties;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.TfsHelper;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using ContentControl = Microsoft.Office.Interop.Word.ContentControl;


namespace AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects
{
    using System.Runtime.InteropServices;
    using Common.Helper;
    using Common.ImageComposer;
    using System.Globalization;

    /// <summary>
    /// This class implements the IField interface for a single word table cell.
    /// </summary>
    public class WordTableField : IField
    {
        private static readonly object HtmlExportLock = new object();

        private readonly IConverter _converter;
        private readonly Range _range;
        private string _valueCache;
        private string _temporaryValue;

        private bool _parseHtmlAsPlaintext;

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTableField"/> class.
        /// </summary>
        /// <param name="range">Range of the document from which to read and write.</param>
        /// <param name="configurationFieldItem">Configuration of the field.</param>
        /// <param name="converter">The converter to use when accessing this fields value.</param>
        /// <param name="cutLastMarker">Defines whether the last character of the range should be excluded.</param>
        public WordTableField(Range range, IConfigurationFieldItem configurationFieldItem, IConverter converter, bool cutLastMarker)
        {
            Guard.ThrowOnArgumentNull(configurationFieldItem, "configurationFieldItem");

            Configuration = configurationFieldItem.Clone();
            _converter = converter;
            _range = range;

            // exclude end-of-cell marker
            if (cutLastMarker)
            {
                if (_range != null)
                    _range.MoveEnd(WdUnits.wdCharacter, -1);
            }

            if (configurationFieldItem.WordBookmark != null && configurationFieldItem.WordBookmark != " ")
            {
                if (_range != null && _range.Text != null)
                {
                    //If the Text of the range is empty, add some non blank space to help word with its Bookmarks
                    if (string.IsNullOrWhiteSpace(_range.Text))
                    {
                        _range.Text = "&nbsp;";
                    }
                }
                if (_range != null && !_range.Bookmarks.Exists(configurationFieldItem.WordBookmark))
                {
                    _range.Bookmarks.Add(configurationFieldItem.WordBookmark);
                }
            }
        }

        #endregion

        #region IField Members

        /// <summary>
        /// Gets the configuration used to create this field and that defines its behavior.
        /// </summary>
        public IConfigurationFieldItem Configuration
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets a list of allowed values for this field. The values are not explicitly checked but used to populate a word dropdown list.
        /// </summary>
        public IList<string> AllowedValues
        {
            get
            {
                return new List<string>();
            }
            set
            {
                SetDropdownValues(value);
            }
        }
        /// <summary>
        /// Gets or sets if the html formating of the field should be ignored.
        /// </summary>
        public bool ParseHtmlAsPlaintext
        {
            get
            {
                return _parseHtmlAsPlaintext;
            }
            set
            {
                _parseHtmlAsPlaintext = value;
            }
        }

        /// <summary>
        /// Gets whether this field is editable
        /// </summary>
        public bool IsEditable
        {
            get
            {
                return true;
            }
        }

        /// <summary>
        /// Gets the reference name of the field.
        /// </summary>
        public string ReferenceName
        {
            get
            {
                return Configuration.ReferenceFieldName;
            }
        }

        /// <summary>
        /// Gets the name of the field.
        /// </summary>
        public string Name
        {
            get
            {
                throw new NotSupportedException();
            }
        }

        /// <summary>
        /// Gets or sets a hyperlink. This is used only in word to create links to TFS web access.
        /// </summary>
        public string Hyperlink
        {
            get
            {
                return string.Empty;
            }

            set
            {
                if (_range != null)
                {
                    _range.Hyperlinks.Add(_range, value);
                }
            }
        }

        /// <summary>
        /// The value of this field. Note that this value is cached and the cache
        /// is only invalidated when setting the value or the special value.
        /// </summary>
        public virtual string Value
        {
            get
            {
                if (_valueCache == null)
                {
                    _valueCache = GetValue();
                }
                return _valueCache;
            }

            set
            {
                // invalidate cache
                _valueCache = null;

                if (Configuration.HandleAsDocument)
                {
                    // store the value for sure if the attached document doesn't exist in opposite adapter work item (TFS)
                    _temporaryValue = value;
                    return;
                }

                SetValue(value, Configuration.FieldValueType);
            }
        }

        /// <summary>
        /// Gets or sets the value of the field without converting it to string.
        /// </summary>
        public object OriginalValue {
            get
            {
                return Value;
            }
        }

        /// <summary>
        /// Select field.
        /// </summary>
        public void NavigateTo()
        {
            if (_range != null)
            {
                _range.Select();
            }
        }

        /// <summary>
        /// Gets or sets the micro document.
        /// </summary>
        public virtual object MicroDocument
        {
            get
            {
                // nothing to do if document handling is turned off or there is no ole element for OleOnDemand
                if (_range == null ||
                    Configuration.HandleAsDocument == false ||
                    (Configuration.HandleAsDocumentMode == HandleAsDocumentType.OleOnDemand && WordSyncHelper.ContainsRangeOleObject(_range) == false))
                {
                    return null;
                }

                var tempPath = Path.Combine(TempFolder.CreateNewTempFolder("WordTableField"), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempPath);

                var file = Path.Combine(tempPath, $"{ReferenceName}.docx");
                _range.ExportFragment(file, WdSaveFormat.wdFormatDocumentDefault);

                return file;
            }

            set
            {
                // if HandleAsDocument flag then load Word document with the field name into range.
                if (!Configuration.HandleAsDocument)
                {
                    return;
                }

                var file = value as string;
                if (string.IsNullOrEmpty(file) == false && File.Exists(file))
                {
                    if (_range != null)
                    {
                        // invalidate cache
                        _valueCache = null;
                        _range.ImportFragment(file);
                    }

                    File.Delete(file);
                    SyncServiceTrace.D("WI?W - Importing fragment {0}", value);
                }
                else
                {
                    // if the file doesn't exist then set the value in a standard way
                    SetValue(_temporaryValue, Configuration.FieldValueType);
                }
            }
        }

        /// <summary>
        /// Gets flag saying whether ole object exists
        /// </summary>
        public bool ContainsOleObject
        {
            get
            {
                return WordSyncHelper.ContainsRangeOleObject(_range);
            }
        }

        /// <summary>
        /// Compares the value of this field to the value of another field.
        /// </summary>
        /// <param name="value">The value to compare to.</param>
        /// <param name="ignoreFormatting">Sets whether formatting is ignored when comparing html fields</param>
        /// <returns>True if the values are equal, False if not.</returns>
        public bool CompareValue(string value, bool ignoreFormatting)
        {
            throw new NotSupportedException("The word table field does not support comparision. Use the tfs field implementation instead.");
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Sets a list of allowed values for this dropdown field. Calling this method on fields not configured as DropDownList has no effect.
        /// </summary>
        /// <param name="allowedValues">List of allowed values</param>
        private void SetDropdownValues(IEnumerable<string> allowedValues)
        {
            if (Configuration.FieldValueType != FieldValueType.DropDownList)
            {
                return;
            }

            ContentControl dropDown;
            var contentControlsCache = _range.ContentControls;

            // Get existing dropdown list or create a new one
            if (contentControlsCache.Count != 1 || contentControlsCache[1].Type != WdContentControlType.wdContentControlDropdownList)
            {
                foreach (var contentControl in contentControlsCache.Cast<ContentControl>())
                {
                    contentControl.Delete(true);
                }

                // Include End-Of-Cell-Marker or else the dropdown does appear in the document but not the object model.
                _range.MoveEnd(WdUnits.wdCharacter, 1);
                dropDown = contentControlsCache.Add(WdContentControlType.wdContentControlDropdownList);
                _range.MoveEnd(WdUnits.wdCharacter, -1);
            }
            else
            {
                dropDown = contentControlsCache[1];
            }

            // Add items
            var listItemsCache = dropDown.DropdownListEntries;
            listItemsCache.Clear();
            foreach (var allowedValue in allowedValues)
            {
                if (allowedValue.Equals(string.Empty))
                {
                    listItemsCache.Add(" ", allowedValue);
                }
                else
                {
                    listItemsCache.Add(allowedValue, allowedValue);
                }
            }
        }

        /// <summary>
        /// Gets the text of associated range.
        /// </summary>
        /// <returns>Range stream.</returns>
        private string GetValueAsPlainText()
        {
            string text;
            if (_range == null || _range.Text == null)
            {
                text = string.Empty;
            }
            else if (_range.ContentControls.Count == 1 && _range.ContentControls[1].ShowingPlaceholderText)
            {
                text = string.Empty;
            }
            else
            {
                var extractedText = _range.Text;
                text = extractedText.Replace("\r", "\n").Replace("\v", "\n");
            }

            return _converter == null ? text : _converter.Convert(text, Direction.OtherToTfs);
        }

        /// <summary>
        /// Gets the range as html by copying it to the clipboard and setting the html format for clipboard.
        /// </summary>
        /// <returns>HTML stream.</returns>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private string GetValueAsHtml()
        {
            // Make sure that when reading from multiple threads, the old values and
            // images are read before they are overwritten by a new read.
            lock (HtmlExportLock)
            {
                try
                {
                    if (_range == null)
                    {
                        return string.Empty;
                    }

                    var helper = new ClipboardHelper(_range);

                    // apply workaround for when only an inline shape is in the field and no text. Word will not export the embedded element if there is no other text.
                    if (string.IsNullOrEmpty(helper.Html))
                    {
                        if (ContainsShapes && Configuration.ShapeOnlyWorkaroundMode == ShapeOnlyWorkaroundMode.AddSpace)
                        {
                            _range.InsertAfter(Resources.ShapeOnlyAutoText);
                            helper = new ClipboardHelper(_range);
                        }
                    }

                    // Generate the file name --> FieldName.Index.Extension --> example AIT.Common.Description.1.png
                    var html = helper.Html;
                    using (var htmlHelper = new HtmlHelper(html, false))
                    {
                        foreach (var image in htmlHelper.Images)
                        {
                            image.Name = Guid.NewGuid().ToString() + image.LocalFileInfo.Extension;
                            html = html.Replace(image.Source, image.LocalFileInfo.FullName);
                        }
                    }

                    /////////////////////////////////////////////////////////////////////////////////////
                    //if (scaleRatio != null)
                    //{
                    //    // Images in word are displayed 20% smaller than they actually are.
                    //    // Thus we have to shrink the images to 80% in order to display them correctly in the TFS fields.
                    //    // On the other hand, when copied into word, image sizes are increased by 25%.
                    //    // Therefore they are again displayed correctly in word when we download them from the TFS.
                    //    html = HtmlHelper.ShrinkImages(html, scaleRatio);
                    //    this.RescalePictures(scaleRatio);
                    //}
                    /////////////////////////////////////////////////////////////////////////////////////

                    return html;
                }
                catch (Exception ex)
                {
                    SyncServiceTrace.LogException(ex);
                    return GetValueAsPlainText();
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether this field contains inline shapes like pictures and smart arts.
        /// </summary>
        /// <value>
        /// <c>true</c> if the field contains inline shapes; otherwise, <c>false</c>.
        /// </value>
        public bool ContainsShapes
        {
            get
            {
                return _range != null && (_range.ShapeRange.Count > 0 || _range.InlineShapes.Count > 0);
            }
        }

        /////////////////////////////////////////////////////////////////////////////////////
        ///// <summary>
        ///// Scale Pictures Up
        ///// </summary>
        //private void ScalePicturesUp()
        //{
        //    foreach (InlineShape shape in _range.InlineShapes)
        //    {
        //        shape.ScaleHeight = 100;
        //        shape.ScaleWidth = 100;
        //    }
        //}

        ///// <summary>
        ///// Rescales the pictures.
        ///// </summary>
        ///// <param name="scaleRatio">The scale ratio.</param>
        //private void RescalePictures(Dictionary<string, float> scaleRatio)
        //{
        //    foreach (InlineShape shape in _range.InlineShapes)
        //    {
        //        if (scaleRatio.ContainsKey(shape.AlternativeText))
        //        {
        //            shape.ScaleHeight = scaleRatio[shape.AlternativeText] * 100;
        //            shape.ScaleWidth = scaleRatio[shape.AlternativeText] * 100;
        //        }
        //    }
        //}

        ///// <summary>
        ///// Calculates for each image a scaling factor to present in properly in Word.
        ///// </summary>
        ///// <returns>
        ///// Dictionary, which assigns to an immage ID the corresponding scaling factor.
        ///// </returns>
        //private Dictionary<String, float> CalculateScaleRatio()
        //{
        //    var scale = new Dictionary<String, float>();
        //    foreach (InlineShape shape in _range.InlineShapes)
        //    {
        //        //Generate a Guid
        //        shape.AlternativeText = Guid.NewGuid().ToString();

        //        //initialize with 1 as fallback;

        //        var calculatedScaleRatio = (float)1.0;

        //        //take the height
        //        if (Math.Abs(shape.ScaleHeight) > 0.00001)
        //        {
        //            calculatedScaleRatio = shape.ScaleHeight / 100;
        //        }

        //        //take the width if is still 1.0
        //        if (Math.Abs(shape.ScaleHeight) > 0.00001)
        //        {
        //            calculatedScaleRatio = shape.ScaleHeight / 100;
        //        }


        //        scale.Add(shape.AlternativeText, calculatedScaleRatio);
        //    }
        //    return scale;
        //}
        /////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Gets a value indicating whether this field contains downscaled pictures.
        /// </summary>
        /// <value>
        /// <c>true</c> if the field contains downscaled pictures; otherwise, <c>false</c>.
        /// </value>
        public bool IsContainsScaledPictures
        {
            get
            {
                if (_range != null)
                {
                    var shapes = _range.InlineShapes.Cast<InlineShape>();
                    foreach (var shape in shapes)
                    {
                        try
                        {
                            if (shape.ScaleHeight < 100 || shape.ScaleWidth < 100)
                            {
                                return true;
                            }
                        }
                        catch (COMException ex)
                        {
                            if (ex.ErrorCode != -2146823595)
                            {
                                throw;
                            }
                            // <hr /> appear as shapes but ScaleWidth and ScaleHeight cannot be set.
                            // The following exception is thrown: "This member cannot be accessed on a horizontal line."
                            SyncServiceTrace.W(string.Format(CultureInfo.CurrentCulture, Resources.Shape_Scale100_Failed, ex.Message));
                        }
                    }
                }
                return false;
            }
        }


        /// <summary>
        /// Gets a value indicating whether the field contains a list item with an end-of-cell marker '\a'.
        /// Word will fail to export this item leaving it unformatted.
        /// </summary>
        /// <value>
        /// <c>true</c> if the range contains an unclosed list; otherwise, <c>false</c>.
        /// </value>
        public bool ContainsUnclosedLists
        {
            get
            {
                return _range.ListParagraphs.Cast<Paragraph>().Any(x => x.Range.Text.Contains("\a"));
            }
        }

        /// <summary>
        /// Gets value for this field.
        /// </summary>
        private string GetValue()
        {
            switch (Configuration.FieldValueType)
            {
                case FieldValueType.PlainText:
                case FieldValueType.BasedOnFieldType:
                    {
                        return GetValueAsPlainText();
                    }

                case FieldValueType.HTML:
                    {
                        if (ParseHtmlAsPlaintext)
                        {
                            return GetHtmlValueAsPlainText();
                        }
                        else
                        {
                            return GetValueAsHtml();
                        }
                    }

                ////case FieldValueType.BasedOnVariable:
                ////return GetValueBasedOnVariable();

                default:
                    {
                        return GetValueAsPlainText();
                    }
            }
        }


        // Get the Value as Html and use a regex
        private string GetHtmlValueAsPlainText()
        {
            var plainTextString = string.Empty;
            if (_range == null || _range.Text == null)
            {
                plainTextString = string.Empty;
            }
            else
            {
                var htmlString = _range.Text.Replace("\r", "\n").Replace("\v", "\n").Trim();
                const string Pattern = "<(.|\n)*?>";
                Regex.Replace(htmlString, Pattern, string.Empty);

                // Remove the rest of any possible html content
                plainTextString = _range.Text.Replace("\r", "\n").Replace("\v", "\n").Trim();
            }

            return _converter == null ? plainTextString : _converter.Convert(plainTextString, Direction.OtherToTfs);
        }

        /// <summary>
        /// Set value for this field.
        /// </summary>
        private void SetValue(string value, FieldValueType fieldValueType)
        {
            if (_range == null)
            {
                return;
            }

            switch (fieldValueType)
            {
                // Set converted value, add hyperlink to range
                case FieldValueType.PlainText:
                    {
                        _range.Text = (_converter == null) ? value : _converter.Convert(value, Direction.TfsToOther);
                        if (!string.IsNullOrEmpty(Hyperlink))
                        {
                            _range.Hyperlinks.Add(_range, Hyperlink);
                        }

                        break;
                    }

                case FieldValueType.BasedOnFieldType:
                    {
                        if (ReferenceName.Equals(FieldReferenceNames.TestSteps))
                        {
                            if (string.IsNullOrEmpty(value) == false)
                            {
                                var template =
                                    _range.Application.ListGalleries[WdListGalleryType.wdNumberGallery].ListTemplates[1];
                                if (template != null)
                                {
                                    _range.ListFormat.ApplyListTemplate(template, false);
                                }
                            }

                            _range.Text = value;
                        }
                        break;
                    }
                case FieldValueType.DropDownList:
                    {
                        var contentControlsCache = _range.ContentControls;

                        if (contentControlsCache.Count != 1)
                        {
                            SyncServiceTrace.E("DropDownList value should be set but the field contains not exactly one FormField.");
                            return;
                        }

                        var dropDown = contentControlsCache[1];
                        var dropDownEntries = dropDown.DropdownListEntries.Cast<ContentControlListEntry>();

                        // if we previously assigned allowed values, check the cache. If not read existing values from dropdown list
                        var matchingEntry = dropDownEntries.FirstOrDefault(x => x.Value == value);
                        if (matchingEntry == null)
                        {
                            SyncServiceTrace.E("DropDownList value should be set but the value is not in the lists of possible values.");
                            return;
                        }

                        matchingEntry.Select();
                        break;
                    }

                case FieldValueType.HTML:
                    {
                        if (string.IsNullOrEmpty(value))
                        {
                            // if the value is empty then do not use clipboard
                            _range.Text = value;
                        }
                        else
                        {
                            // Fix for 15925. Editing html fields using vs or browser will not always surround text in tags.
                            // Word appears to only parse strings like "A&amp;B" correctly when they are surrounded with tags.
                            if (!(value.StartsWith("<", StringComparison.Ordinal) && value.EndsWith(">", StringComparison.Ordinal)))
                            {
                                value = $"<span>{value}</span>";
                                SyncServiceTrace.W("Automatically surrounded text with span tags");
                            }

                            value = HtmlHelper.AddMissingImageSizeAttributes(value);
                            // value = HtmlHelper.ShrinkImagesAllSamePercentage(value, 0.8f);

                            // Get the range to work with
                            Range range;
                            if (!string.IsNullOrEmpty(Configuration.WordBookmark))
                            {
                                // Set the Bookmarks
                                // If the Text of the range is empty, add some non blank space to help word with its Bookmarks
                                if (string.IsNullOrWhiteSpace(_range.Text))
                                {
                                    _range.Text = "&nbsp;";
                                }

                                _range.Bookmarks.Add(Configuration.WordBookmark);
                                range = _range.Bookmarks[Configuration.WordBookmark].Range;
                            }
                            else
                            {
                                range = _range;
                            }

                            // Get the start of actual range to select whole pasted 'thing'
                            var startRange = range.Start;

                            try
                            {
                                WordHelper.PasteSpecial(range, value);
                            }
                            catch (COMException ce)
                            {
                                if (ce.ErrorCode == WordHelper.CommandFailedCode)
                                {
                                    _range.Text = string.Empty;
                                    SyncServiceTrace.I("HTML-content \"{0}\" was effectively empty and therefore it has been replaced by an empty string.", value);
                                }
                                else
                                {
                                    throw;
                                }
                            }

                            // Set the start back to the previous position to set the pasted 'thing' into this range
                            range.Start = startRange;
                            var parser = new PasteStreamParser();
                            parser.ParseAndRepairAfterPaste(value, range);
                        }
                        break;
                    }
            }

            RefreshBookmarks();
        }

        /// <summary>
        /// Refreshes the Bookmarks of the field, this should be used after setting the values
        /// </summary>
        public void RefreshBookmarks()
        {
            if (!string.IsNullOrEmpty(Configuration.WordBookmark))
            {
                _range.Bookmarks.Add(Configuration.WordBookmark);
            }
        }

        #endregion

    }
}
