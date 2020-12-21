#region Usings
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Factory;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using AIT.TFS.SyncService.Common.Helper;
using AIT.TFS.SyncService.Adapter.TFS2012.Properties;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects
{
    /// <summary>
    /// Class represents one field of one <see cref="IWorkItem"/>.
    /// </summary>
    public class TfsField : IField,IComparable
    {
        #region Fields
        private readonly Field _field;
        private HtmlHelper _htmlHelper;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="TfsField"/> class.
        /// </summary>
        /// <param name="workItem">Work item to present one field from.</param>
        /// <param name="configurationFieldItem">Configuration field item.</param>
        internal TfsField(TfsWorkItem workItem, IConfigurationFieldItem configurationFieldItem)
        {
            Guard.ThrowOnArgumentNull(workItem, "workItem");
            Guard.ThrowOnArgumentNull(configurationFieldItem, "configurationFieldItem");

            WorkItem = workItem;
            Configuration = configurationFieldItem;
            _field = workItem.WorkItem.Fields[Configuration.ReferenceFieldName];
        }
        #endregion

        #region Properties

        /// <summary>
        /// Work item to present a field from.
        /// </summary>
        protected TfsWorkItem WorkItem
        {
            get;
            set;
        }

        /// <summary>
        /// Type of the field. See <see cref="FieldValueType"/>.
        /// </summary>
        private FieldValueType ValueType
        {
            get
            {
                return Configuration.FieldValueType;
            }
        }

        /// <summary>
        /// Gets the configuration defining the synchronization behavior of this field
        /// </summary>
        public IConfigurationFieldItem Configuration
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets a list of allowed values for this field. The values are not explicitly checked but used to populate a word dropdown list.
        /// </summary>
        public virtual IList<string> AllowedValues
        {
            get
            {
                var allowedValues = _field.AllowedValues.Cast<string>().ToList();
                // Wenn Feld nicht 'required' ist, wird leerer String hinzugefügt um herauslöschen aus dropdown list zu ermöglichen
                if (!_field.IsRequired)
                {
                    allowedValues.Add(string.Empty);
                }
                return allowedValues;
            }
            set
            {
                // ReadOnly Property in tfs
            }
        }


        private string MicroDocumentFile
        {
            get;
            set;
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
                return _field.Name;
            }
        }


        /// <summary>
        /// Gets or sets a hyperlink. This is used in word only to create links to web access to TFS.
        /// </summary>
        public string Hyperlink
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the value of the field.
        /// </summary>
        public virtual string Value
        {
            get
            {
                return Convert.ToString(_field.Value, CultureInfo.CurrentCulture);
            }
            set
            {
                if (_field.IsEditable)
                {
                    WorkItem.OpenWorkItemForWrite();

                    if (FieldValueType.HTML == ValueType)
                    {
                        _htmlHelper = new HtmlHelper(value, true);

                        // Update image names to "System.Description.0.jpg" etc.
                        for (var index = 0; index < _htmlHelper.Images.Count; index++)
                        {
                            var image = _htmlHelper.Images[index];
                            image.Name = StringHelper.GetImageAttachmentName(ReferenceName, index, image.LocalFileInfo.Extension);
                        }

                        if (!CompareImages(_htmlHelper))
                        {
                            UpdateAllImages();
                        }
                        else
                        {
                            SetHTMLValueWithUpdatedImageSources();
                        }
                    }
                    else
                    {
                        if (_field.IsRequired || value != null)
                        {
                            var convertedValue = Convert.ChangeType(value, _field.FieldDefinition.SystemType, CultureInfo.CurrentCulture);
                            SetValue(convertedValue);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the value of the field without converting it to string.
        /// </summary>
        public object OriginalValue
        {
            get
            {
                return _field.Value;
            }
        }

        /// <summary>
        /// Gets or sets the micro document that exports the field value.
        /// </summary>
        public virtual object MicroDocument
        {
            get
            {
                if (Configuration.HandleAsDocument)
                {
                    // find micro document attachment and save it locally
                    var fileName = StringHelper.GetMicroDocumentName(ReferenceName);
                    var attachment = GetAttachment(fileName);
                    if (attachment != null)
                    {
                        var file = Path.Combine(TempFolder.CreateNewTempFolder("TFSField"), fileName);
                        TfsAttachment.DownloadWebFile(attachment.Uri, WorkItem.WorkItemStore.TeamProjectCollection.Credentials, file);
                        return file;
                    }
                }

                return null;
            }
            set
            {
                if (Configuration.HandleAsDocument)
                {
                    // delete micro document attachment
                    RemoveAttachments(name => StringHelper.GetMicroDocumentName(ReferenceName).Equals(name, StringComparison.OrdinalIgnoreCase));

                    // add new micro document
                    var file = value as string;
                    if (string.IsNullOrEmpty(file) == false && File.Exists(file))
                    {
                        // set new attachment
                        MicroDocumentFile = file;
                        WorkItem.WorkItem.Attachments.Add(new Attachment(file));
                        SyncServiceTrace.D(Resources.AttachingMicroDocument, WorkItem.Id, value);
                    }
                }
            }
        }

        /// <summary>
        /// Gets whether this field is editable
        /// </summary>
        /// <returns>The value of the actual TFS work item field property.</returns>
        public bool IsEditable
        {
            get
            {
                return _field.IsEditable;
            }
        }

        /// <summary>
        /// Gets flag saying whether ole object exists
        /// </summary>
        public bool ContainsOleObject
        {
            get
            {
                throw new NotSupportedException();
            }
        }
        #endregion


        /// <summary>
        /// Refreshes the Bookmarks of the field, this should be used after setting the values
        /// </summary>
        public void RefreshBookmarks()
        {
            throw new NotSupportedException();
        }
        /// <summary>
        /// Navigates the user to this field.
        /// </summary>
        /// <exception cref="System.NotSupportedException">Not supported for this adapter.</exception>
        public void NavigateTo()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Compares the value of this field to another value. This comparison uses the ignoreFormatting
        /// property set in the configuration for this field.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="ignoreFormatting">Sets whether html formatting is ignored.</param>
        /// <returns>True if the values are equal, False if not.</returns>
        public virtual bool CompareValue(string value, bool ignoreFormatting)
        {
            if ((value == null || Value == null) && value != Value) return false;

            if(ignoreFormatting)
            {
                if (ValueType == FieldValueType.HTML)
                {
                    // compare plain text extracted from HTML content
                    using (var htmlHelper = new HtmlHelper(value, false))
                    {
                        if (HtmlHelper.ExtractPlaintext(Value).Equals(HtmlHelper.ExtractPlaintext(htmlHelper.Html)))
                        {
                            return CompareImages(htmlHelper);
                        }
                    }

                    return false;
                }
            }
            foreach(Project project in WorkItem.WorkItemStore.Projects)
            {
                if(Value.StartsWith(project.Name))
                {
                    var valueWithoutProjectName = string.Empty;
                    valueWithoutProjectName = Value;
                    valueWithoutProjectName = valueWithoutProjectName.Remove(0, project.Name.Length);
                    if (Value.Equals(project.Name)) valueWithoutProjectName = "\\";
                    return valueWithoutProjectName == value;
                }
            }
            return Value == value;
        }

        /// <summary>
        /// Update the source attribute of the image tags to point to the attachment URIs and assign the html to the field.
        /// </summary>
        internal void SetHTMLValueWithUpdatedImageSources()
        {
            if (FieldValueType.HTML == ValueType && null != _htmlHelper)
            {
                WorkItem.OpenWorkItemForWrite();

                foreach (HtmlImage htmlImage in _htmlHelper.Images)
                {
                    var attachment = GetAttachment(htmlImage.Name);
                    if (null != attachment)
                    {
                        htmlImage.Uri = attachment.Uri;
                    }
                    else
                    {
                        SyncServiceTrace.E(Resources.CannotUpdateImageSource, WorkItem.WorkItem.Id, htmlImage.Uri, htmlImage.Name);
                    }
                }

                SetValue(_htmlHelper.Html);
                _htmlHelper.Dispose();
                _htmlHelper = null;
            }
        }

        /// <summary>
        /// The value must be first set to TFS WI and then can be invalid characters removed
        /// (It seems that some conversion is done when value is attached to TFS WI).
        /// </summary>
        private void SetValue(object value)
        {
            _field.Value = value;
            _field.Value = RemoveInvalidCharacters(Value);
        }

        /// <summary>
        /// Delete temporary files after save process.
        /// </summary>
        internal void DeleteTempFileAfterSave()
        {
            if (File.Exists(MicroDocumentFile))
            {
                var directory = Path.GetDirectoryName(MicroDocumentFile);
                if (string.IsNullOrEmpty(directory) == false)
                {
                    Directory.Delete(directory, true);
                }
            }
        }

        /// <summary>
        /// Compares all images within the current and the given html stream.
        /// </summary>
        /// <returns>True if the images are identical, otherwise false.</returns>
        private bool CompareImages(HtmlHelper htmlHelper)
        {
            using (var htmlHelperThisField = new HtmlHelper(Value, true))
            {
                if (htmlHelperThisField.Images.Count != htmlHelper.Images.Count)
                {
                    return false;
                }

                for (var i = 0; i < htmlHelperThisField.Images.Count; i++)
                {
                    var image = htmlHelperThisField.Images.ElementAt(i);
                    var wordImage = htmlHelper.Images.ElementAt(i);

                    // word sources are assumed to be local files. Get the attachment from the tfs source
                    var attachment = GetAttachment(image.FileNameQueryParameter);
                    if (null == attachment)
                        return false;

                    // compare local file to attachment
                    using (var attachmentHelper = new TfsAttachment(WorkItem, attachment))
                    {
                        if (attachmentHelper.Compare(wordImage.LocalFileInfo) == false)
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Deletes all attachments that match given criteria
        /// </summary>
        private void RemoveAttachments(Func<string, bool> selector)
        {
            for (var i = WorkItem.WorkItem.Attachments.Count - 1; i >= 0; i--)
            {
                var attachment = WorkItem.WorkItem.Attachments[i];
                if (selector(attachment.Name))
                {
                    WorkItem.WorkItem.Attachments.RemoveAt(i);
                }
            }
        }

        /// <summary>
        /// Drops all existing html images and adds all currently existing.
        /// </summary>
        private void UpdateAllImages()
        {
            RemoveAttachments(name => name.StartsWith(ReferenceName, StringComparison.OrdinalIgnoreCase) && StringHelper.GetMicroDocumentName(ReferenceName).Equals(name, StringComparison.OrdinalIgnoreCase) == false);

            if (null != _htmlHelper)
            {
                foreach(var image in _htmlHelper.Images)
                {
                    var attachment = new Attachment(image.LocalFileInfo.FullName, string.Empty);
                    WorkItem.WorkItem.Attachments.Add(attachment);
                }
            }
        }

        /// <summary>
        /// Get attachment by name.
        /// </summary>
        protected virtual Attachment GetAttachment(string name)
        {
            var attachments = WorkItem.WorkItem.Attachments.Cast<Attachment>();
            return attachments.FirstOrDefault(attachment => attachment.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Remove invalid characters from text.
        /// </summary>
        private static string RemoveInvalidCharacters(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var validXml = new StringBuilder(value.Length, value.Length);
            var charArray = value.ToCharArray();

            for (int i = 0; i < charArray.Length; i++)
            {
                char current = charArray[i];
                if ((current == 0x9) ||
                    (current == 0xA) ||
                    (current == 0xD) ||
                    ((current >= 0x20) && (current <= 0xD7FF)) ||
                    ((current >= 0xE000) && (current <= 0xFFFD)))
                    validXml.Append(current);
            }

            return validXml.ToString();
        }

        public int CompareTo(object obj)
        {
            var otherField = obj as TfsField;
            return ReferenceName.CompareTo(otherField.ReferenceName);
        }

        /// <summary>
        /// Gets a value indicating whether this field contains inline shapes like pictures and smart arts.
        /// </summary>
        public bool ContainsShapes
        {
            get
            {
                throw new NotSupportedException();
            }
        }
    }
}