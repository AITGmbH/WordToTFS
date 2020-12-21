using System.Collections.Generic;

namespace AIT.TFS.SyncService.Service.Utils
{

    /// <summary>
    /// Class that helps to parse information that is separated by a CSV
    /// </summary>
   public static class CsvParser
    {

        /// <summary>
        /// Parse a CSV Value and return a list of strings
        /// </summary>
        /// <param name="csvList">A string that contans csv separated data</param>
        /// <returns></returns>
        public static List<string> ParseCsv(string csvList)
        {
            var returnList = new List<string>();
            if (csvList.Contains(","))
            {
                string[] values = csvList.Split(',');
                foreach (string singleValue in values)
                {
                    var trimmedValue = singleValue.Trim();
                    returnList.Add(trimmedValue);
                }
            }
            else
            {
                returnList.Add(csvList);
            }
            return returnList;
        }

    }
}
