using AIT.TFS.SyncService.Contracts.Configuration;
using System.Collections.Generic;

namespace AIT.TFS.SyncService.Contracts.WorkItemObjects
{
    /// <summary>
    /// Interface defines functionality of a field in one work item.
    /// </summary>
    public interface IField
    {
        /// <summary>
        /// Gets the reference name of the field.
        /// </summary>
        string ReferenceName { get; }

        /// <summary>
        /// Gets the name of the field.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets or sets the value of the field.
        /// </summary>
        string Value { get; set; }

        /// <summary>
        /// Gets or sets the value of the field without converting it to string.
        /// </summary>
        object OriginalValue { get; }

        /// <summary>
        /// Gets or sets a hyperlink. This is used in word only to create links to web access to TFS.
        /// </summary>
        string Hyperlink { get; set; }

        /// <summary>
        /// Navigates the user to this field.
        /// </summary>
        void NavigateTo();

        /// <summary>
        /// Gets or sets the micro document that exports the field value.
        /// </summary>
        object MicroDocument { get; set; }

        /// <summary>
        /// Gets flag saying whether ole object exists
        /// </summary>
        bool ContainsOleObject { get; }

        /// <summary>
        /// Gets a value indicating whether this field contains inline shapes like pictures and smart arts.
        /// </summary>
        bool ContainsShapes { get; }

        /// <summary>
        /// Compares the value of this field to the value of another field.
        /// </summary>
        /// <param name="value">The value to compare to.</param>
        /// <param name="ignoreFormatting">Sets whether formatting is ignored when comparing html fields</param>
        /// <returns>True if the values are equal, False if not.</returns>
        bool CompareValue(string value, bool ignoreFormatting);

        /// <summary>
        /// Gets whether this field is editable.
        /// </summary>
        bool IsEditable { get; }

        /// <summary>
        /// Gets the configuration used to create this field and that defines its behavior.
        /// </summary>
        IConfigurationFieldItem Configuration { get; }

        /// <summary>
        ///  Gets or sets a list of allowed values for this field. The values are not explicitly checked but used to populate a word dropdown list.
        /// </summary>
        IList<string> AllowedValues
        {
            get;
            set;
        }

        /// <summary>
        /// Refreshes the Bookmarks of the field, this should be used after setting the values
        /// </summary>
        void RefreshBookmarks();
    }
}