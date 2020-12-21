using AIT.TFS.SyncService.Adapter.Word2007.WorkItemObjects;
using AIT.TFS.SyncService.Model;
using AIT.TFS.SyncService.Model.WindowModel;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using Microsoft.Office.Interop.Word;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Windows;
using System.Windows.Controls;
using System;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for GetWorkItemsPanelView.
    /// </summary>
    public partial class SynchronizationStatePanelView : IHostPaneControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetWorkItemsPanelView"/> class.
        /// </summary>
        public SynchronizationStatePanelView()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Gets or sets the ViewModel for this View.
        /// </summary>
        public WorkItemOverviewPanelViewModel Model
        {
            get
            {
                return DataContext as WorkItemOverviewPanelViewModel;
            }
            set
            {
                DataContext = value;
                Load();
            }
        }

        /// <summary>
        /// Gets or sets Word Document instance.
        /// </summary>
        public Document AttachedDocument
        {
            get { return Model.WordDocument; }
        }

        /// <summary>
        /// The method is called after the visibility of the containing panel has changed.
        /// </summary>
        /// <param name="isVisible"><c>true</c>if the panel containing this view is shown. False otherwise.</param>
        public void VisibilityChanged(bool isVisible)
        {
            if (isVisible == false)
            {
                TreeViewPopup.IsOpen = false;
                Model.Cancel();
            }
        }

        private void SelectQuery(object query)
        {
            var queryItem = query as QueryItem;
            var queryFolder = query as QueryFolder;

            if (queryFolder != null)
            {
                return;
            }

            TreeViewPopup.IsOpen = false;
            QueryBox.Items.Clear();

            if (queryItem == null)
            {
                Model.SelectedQuery = null;
                QueryBox.Items.Add("All Work Items in the document");
            }
            else
            {
                Model.SelectedQuery = queryItem;
                QueryBox.Items.Add(Model.SelectedQuery.Name);
            }

            QueryBox.SelectedItem = QueryBox.Items[0];
            Model.FindWorkItems();

        }

        private void Find_Click(object sender, RoutedEventArgs e)
        {
            Model.FindWorkItems();
        }

        private void Load()
        {
            Model.UseLinkedWorkItems = Model.QueryConfiguration.UseLinkedWorkItems;
            Dispatcher.Invoke(new Action(() => SelectQuery(Model.SelectedQuery)));
        }

        /// <summary>
        /// Show or hide query tree popup.
        /// </summary>
        private void QueryBox_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            e.Handled = true;
            TreeViewPopup.IsOpen = !TreeViewPopup.IsOpen;
        }

        /// <summary>
        /// Select a query and load it if it is different to the current shown query
        /// </summary>
        private void QueryTree_OnSelected(object sender, RoutedEventArgs e)
        {
            var query = ((TreeViewItem) e.OriginalSource).DataContext as QueryItem;

            if (query != null)
            {
                if (Model.SelectedQuery != null && Model.SelectedQuery.Id == query.Id)
                {
                    return;
                }
            }

            SelectQuery(query);
        }

        /// <summary>
        /// Navigate to the table of the selected work item.
        /// </summary>
        private void FoundList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 0)
                return;

            var itemViewModel = e.AddedItems[0] as DataItemModel<SynchronizedWorkItemViewModel>;
            if (itemViewModel == null)
                return;

            if(itemViewModel.Item.WordItem is WordTableWorkItem)
                (itemViewModel.Item.WordItem as WordTableWorkItem).Table.Select();
        }
    }
}
