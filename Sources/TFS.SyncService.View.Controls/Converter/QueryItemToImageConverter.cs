#region Usings
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Converter
{
    /// <summary>
    /// Converter that takes a Node type and an isExpanded state and converts it into an image representation.
    /// </summary>
    public sealed class QueryItemToImageConverter : IMultiValueConverter
    {
        private const string ResourcePath = "pack://application:,,,/{0};component/Resources/";
        private const string TeamExplorerFlatList = "TeamExplorerFlatList.png";
        private const string TeamExplorerMyQueries = "TeamExplorerMyQueries.png";
        private const string TeamExplorerDirectLink = "TeamExplorerDirectLink.png";
        private const string TeamExplorerFolderCollapsed = "TeamExplorerFolderCollapsed.png";
        private const string TeamExplorerFolderExpanded = "TeamExplorerFolderExpanded.png";
        private const string TeamExplorerNoWorkItems = "TeamExplorerNoWorkItems.png";
        private const string TeamExplorerTeamQueries = "TeamExplorerTeamQueries.png";
        private const string TeamExplorerTree = "TeamExplorerTree.png";
        private const string Error = "Error.png";

        /// <summary>
        /// Converts the node type into a bitmap representation.
        /// </summary>
        /// <param name="values">An array of length 2 where the first entry is a QueryItem and the second value a boolean indicating whether the QueryItem is expanded or not.</param>
        /// <param name="targetType">Not used.</param>
        /// <param name="parameter">Not used.</param>
        /// <param name="culture">Not used.</param>
        /// <returns></returns>
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 2) return DependencyProperty.UnsetValue;

            var folder = values[0] as QueryFolder;
            var definition = values[0] as QueryDefinition;

            var resourcePath = string.Format(CultureInfo.InvariantCulture, ResourcePath, GetType().Assembly.GetName().Name);
            var imagePath = resourcePath + TeamExplorerFolderCollapsed;

            var isExpandend = values[1] is bool && (bool)values[1];

            if (definition != null)
            {
                switch (definition.QueryType)
                {
                    case QueryType.Invalid:
                        imagePath = resourcePath + TeamExplorerNoWorkItems;
                        break;
                    case QueryType.List:
                        imagePath = resourcePath + TeamExplorerFlatList;
                        break;
                    case QueryType.OneHop:
                        imagePath = resourcePath + TeamExplorerDirectLink;
                        break;
                    case QueryType.Tree:
                        imagePath = resourcePath + TeamExplorerTree;
                        break;
                    default:
                        imagePath = resourcePath + Error;
                        break;
                }
            }
            else if (folder != null)
            {
                if (folder.Parent is QueryHierarchy)
                {
                    // Our parent is the hierarchy so we are either My Queries or Team Queries
                    imagePath = folder.IsPersonal ? resourcePath + TeamExplorerMyQueries : resourcePath + TeamExplorerTeamQueries;
                }
                else
                {
                    imagePath = isExpandend ? resourcePath + TeamExplorerFolderExpanded : resourcePath + TeamExplorerFolderCollapsed;
                }
            }

            try
            {
                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.UriSource = new Uri(imagePath, UriKind.RelativeOrAbsolute);
                bitmapImage.EndInit();
                return bitmapImage;
            }
            catch (InvalidOperationException)
            {
                return DependencyProperty.UnsetValue;
            }
        }

        /// <summary>
        /// Not implemented.
        /// </summary>
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}