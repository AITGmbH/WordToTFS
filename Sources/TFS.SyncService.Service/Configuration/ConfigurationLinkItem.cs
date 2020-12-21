using System;
using System.Linq;
using System.Text.RegularExpressions;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Contracts.Exceptions;
using System.Collections.Generic;
using System.Globalization;
using AIT.TFS.SyncService.Common.Helper;

namespace AIT.TFS.SyncService.Service.Configuration
{
    /// <summary>
    /// Class that represents a single work item field.
    /// </summary>
    internal class ConfigurationLinkItem : IConfigurationLinkItem
    {
        #region Private consts

        // Default values used when the xml attributes are missing.
        private const string DefaultLinkSeparator = ", ";
        private const string DefaultLinkFormat = "{System.Id}";

        #endregion Private consts

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationFieldItem"/> class.
        /// </summary>
        /// <param name="linkValueType">Determines whether the exported value should be plain text or html.</param>
        /// <param name="linkedWorkItemTypes">Determines a comma separated list of Work Item Types that can be used to filter the links in addition to the <see cref="linkValueType"/>.</param>
        /// <param name="direction">Determines the copy direction of the field.</param>
        /// <param name="rowIndex">Table row index.</param>
        /// <param name="colIndex">Table column index.</param>
        /// <param name="overwrite">Overwrite links</param>
        /// <param name="linkSeparator">A text placed in between links to different work items</param>
        /// <param name="linkFormat">A text with named format sequences that are to be replaced by values of referenced properties of the linked work item</param>
        /// <param name="automaticLinkWorkItemType">If given, the work item is automatically linked to the previous (in Word) Work Item of this type</param>
        /// <param name="automaticWorkItemSubtypeField">If given, the work item is automatically linked to the previous (in Word) Work Item of type WorkItemType if its field has the value SubTypeValue</param>
        /// <param name="automaticLinkWorkItemSubtypeValue">If given, the work item is automatically linked to the previous (in Word) Work Item of type WorkItemType if its field has the value SubTypeValue</param>
        /// <param name="automaticLinkSuppressWarning">Sets whether to show an error if no automatic link target was found.</param>
        public ConfigurationLinkItem(string linkValueType, string linkedWorkItemTypes, Direction direction, int rowIndex, int colIndex, bool overwrite, string linkSeparator, string linkFormat, string automaticLinkWorkItemType, string automaticWorkItemSubtypeField, string automaticLinkWorkItemSubtypeValue, bool automaticLinkSuppressWarning)
        {
            LinkValueType = linkValueType;
            LinkedWorkItemTypes = linkedWorkItemTypes;
            Direction = direction;
            ColIndex = colIndex;
            RowIndex = rowIndex;
            Overwrite = overwrite;
            LinkSeparator = string.IsNullOrEmpty(linkSeparator)
                                ? DefaultLinkSeparator
                                : linkSeparator.Replace("\\t", "\t").Replace("\\n", "\n").Replace("\\r", "\r");
            LinkFormat = string.IsNullOrEmpty(linkFormat)
                             ? DefaultLinkFormat
                             : linkFormat;

            AutomaticLinkWorkItemType = automaticLinkWorkItemType;
            AutomaticLinkWorkItemSubtypeField = automaticWorkItemSubtypeField;
            AutomaticLinkWorkItemSubtypeValue = automaticLinkWorkItemSubtypeValue;
            AutomaticLinkSuppressWarnings = automaticLinkSuppressWarning;
        }

        #endregion Constructors

        #region Implementation of IConfigurationLinkItem

        /// <summary>
        /// The type of the link
        /// </summary>
        public string LinkValueType { get; private set; }

        /// <summary>
        /// Determines the copy direction of the field.
        /// </summary>
        public Direction Direction { get; private set; }

        /// <summary>
        /// Gets the table column index.
        /// </summary>
        public int ColIndex { get; private set; }

        /// <summary>
        /// Gets the table column index.
        /// </summary>
        public int RowIndex { get; private set; }

        /// <summary>
        /// Gets the value to indicate if empty cell deleting links means.
        /// </summary>
        public bool Overwrite { get; private set; }

        /// <summary>
        /// Gets the string used to separate links within the same cell
        /// </summary>
        public string LinkSeparator { get; private set; }

        /// <summary>
        /// Gets the format string used to format output to represent a link
        /// </summary>
        public string LinkFormat { get; private set; }

        /// <summary>
        /// Gets the type of work item to which to link automatically when publishing
        /// </summary>
        public string AutomaticLinkWorkItemType
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets an optional field used to select the automatically linked work item
        /// </summary>
        public string AutomaticLinkWorkItemSubtypeField
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the required value of the optional subtype field
        /// </summary>
        public string AutomaticLinkWorkItemSubtypeValue
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets whether to suppress warnings when no automatic link target was found
        /// </summary>
        public bool AutomaticLinkSuppressWarnings
        {
            get;
            private set;
        }

        /// <summary>
        /// Formats a work item by the set <see cref="LinkFormat"/>
        /// </summary>
        /// <param name="workItem">Work item to format.</param>
        /// <returns>string representing the work item. If an error occurred, this is only the work item id</returns>
        public string Format(IWorkItem workItem)
        {
            var formattedworkItem = WorkItemStringFormatter.GetWorkItemFomatted(workItem, LinkFormat);
            return formattedworkItem;
        }

        /// <summary>
        /// returns all fields that are necessary to format links
        /// </summary>
        public IEnumerable<string> GetLinkFormatRequiredFields()
        {
            List<string> fields = new List<string>();
            foreach (var match in Regex.Matches(LinkFormat, "{[^}]+}", RegexOptions.CultureInvariant))
            {
                fields.Add(match.ToString().Trim(new[] { '}', '{' }));
            }
            return fields.Distinct();
        }

        /// <summary>
        /// Extracts the work item id from a formatted string. Uses the set <see cref="LinkFormat"/> to find id.
        /// </summary>
        /// <param name="formattedString">string representing the work item or only an id.</param>
        /// <returns>Work item id or Null if the string is not correctly formatted and not an id.</returns>
        /// <exception cref="ConfigurationException">If the id could not be extracted from the formatted string</exception>
        /// <exception cref="OverflowException">When the value in the id field could not be casted into an int</exception>
        public int? GetWorkItemId(string formattedString)
        {
            if (string.IsNullOrEmpty(formattedString))
                return null;


            // Create regexp to find id in a formatted link output
            // Create capture groups: "{FieldName}" => "(?<FieldName>.*)"
            // and try to find the capture group "System.Id" (System_Id
            // because capture group names cannot contain interpunctuation
            try
            {
                // 'Replace("–", "-")' => Replaces the word long dash
                MatchEvaluator capture =
                    x => "(?<" + x.ToString().Substring(1, x.Length - 2).Replace(".", "_").Replace("–", "-") + @">.*)";
                var reverse = Regex.Replace(LinkFormat, "{[^}]+}", capture);

                var match = Regex.Match(formattedString, reverse, RegexOptions.ExplicitCapture);
                var systemId = match.Groups["System_Id"];
                var captures = systemId.Captures;
                var value = captures.Count == 1 ? captures[0].Value : formattedString;

                // In some cases 'value' is number and text behind this number (if the delimiter symbol occors many times)
                // We implements on this place a little extension
                if (!string.IsNullOrEmpty(value))
                {
                    // Trim non digit chars on the right side
                    while (!string.IsNullOrEmpty(value) && !char.IsDigit(value[value.Length - 1]))
                        value = value.Substring(0, value.Length - 1);
                    // Trim non digit chars on the left side
                    while (!string.IsNullOrEmpty(value) && !char.IsDigit(value[value.Length - 1]))
                        value = value.Substring(0, value.Length - 1);
                }

                return int.Parse(value, CultureInfo.InvariantCulture);
            }
            catch (ArgumentException argumentException) { throw new ConfigurationException($"Cannot apply regular expression to {formattedString}", argumentException); }
            catch (FormatException formatException) { throw new ConfigurationException("The id is not numeric or the link format is ambigious.", formatException); }
        }

        public IEnumerable<string> GetLinkedWorkItemTypes()
        {
            var linkedWorkItemTypes = new List<string>();
            if (LinkedWorkItemTypes != null)
            {
                // split by comma
                // after that trim each element of the string array that results from the plit operation before
                linkedWorkItemTypes = LinkedWorkItemTypes.Split(',').Select(s => s.Trim()).ToList();
            }
            return linkedWorkItemTypes;
        }

        /// <summary>
        /// Gets the ...
        /// </summary>
        public string LinkedWorkItemTypes { get; private set; }

        #endregion Implementation of IConfigurationLinkItem
    }
}