#region Usings
using System;
using System.Text.RegularExpressions;
using AIT.TFS.SyncService.Contracts.Enums;
using AIT.TFS.SyncService.Contracts.Exceptions;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
#endregion

namespace AIT.TFS.SyncService.Common.Helper
{
    /// <summary>
    /// This class helps to format work item informations
    /// </summary>
    public class WorkItemStringFormatter
    {
        /// <summary>
        /// Gives a formatted string for a work item based on the link format and using the configuration of the work item.
        /// </summary>
        /// <param name="workItem">The work item which should be formatted.</param>
        /// <param name="linkFormat">The link formatting (e.g. "{System.Id} - {System.Description} - {Customer.CustomField} "</param>
        /// <returns></returns>
        public static string GetWorkItemFomatted(IWorkItem workItem, string linkFormat)
        {
            Func<string, IWorkItem, string> replacementSearcher = (refNameWithAddition, funcWorkItem) =>
            {
                var plainText = string.Empty;
                var refName = refNameWithAddition.ToString().Trim(new[] { '}', '{' });

                if (workItem.Fields.Contains(refName))
                {
                    plainText = workItem.Fields[refName].Value;
                    if (workItem.Fields[refName].Configuration.FieldValueType == FieldValueType.HTML)
                    {
                        var htmlText = workItem.Fields[refName].Value;
                        plainText = HtmlHelper.ExtractPlaintext(htmlText);
                    }
                }
                else
                {
                    throw new ConfigurationException($"Field ({refName}) is not configured in mapping section for work item type ({workItem.WorkItemType}).");
                }

                return plainText;
            };

            var replacedPlainTxtOnly = Regex.Replace(linkFormat, "{[^}]+}", x => replacementSearcher(x.ToString(), workItem));

            return replacedPlainTxtOnly;
        }
    }
}
