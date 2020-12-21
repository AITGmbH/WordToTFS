using System.Globalization;

namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects
{

    /// <summary>
    /// Helper class to handle the strings.
    /// </summary>
    public static class StringHelper
    {
        /// <summary>
        /// Gets micro document file name.
        /// </summary>
        /// <param name="fieldName">Field name to which the document belongs.</param>
        /// <returns></returns>
        public static string GetMicroDocumentName(string fieldName)
        {
            return $"{fieldName}.docx";
        }

        /// <summary>
        /// Gets attachment name for image embedded in html.
        /// </summary>
        public static string GetImageAttachmentName(string referenceFieldName, int index, string fileExtension)
        {
            return $"{referenceFieldName}.{index}{fileExtension}";
        }
    }
}
