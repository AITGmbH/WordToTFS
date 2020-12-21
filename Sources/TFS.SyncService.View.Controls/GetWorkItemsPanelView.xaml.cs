using AIT.TFS.SyncService.Model;
using AIT.TFS.SyncService.Model.WindowModel;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using Microsoft.Office.Interop.Word;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Input;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for GetWorkItemsPanelView.
    /// </summary>
    public partial class GetWorkItemsPanelView : IHostPaneControl
    {
        private DispatcherTimer _longOperationTimer;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetWorkItemsPanelView"/> class.
        /// </summary>
        public GetWorkItemsPanelView()
        {
            InitializeComponent();
        }

        private void SetTreeViewExpandationDefault(bool collapsTree)
        {
            // Info: setter defined in xaml
            var expandationSetter = QueryTree.ItemContainerStyle.Setters.OfType<Setter>().FirstOrDefault(x => x.Property.Name == "IsExpanded");
            if (expandationSetter != null)
            {
                expandationSetter.Value = !collapsTree;
            }
        }

        /// <summary>
        /// Gets or sets the ViewModel for this View.
        /// </summary>
        public GetWorkItemsPanelViewModel Model
        {
            get
            {
                return DataContext as GetWorkItemsPanelViewModel;
            }
            set
            {
                DataContext = value;
                SetTreeViewExpandationDefault(value.CollapseQueryTree);
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
            Model.Cancel();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            Model.Import();
        }

        private void CreateHyperlink_Click(object sender, RoutedEventArgs e)
        {
            Model.CreateHyperlink();
        }

        private void OnKeyDownHandler(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return && !string.IsNullOrWhiteSpace(IdsTextBox.Text))
            {
                Find_Click(sender, e);
            }
        }

        private void Find_Click(object sender, RoutedEventArgs e)
        {
            Model.FindWorkItems();
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            SelectAllFoundedWorkItems(true);
            Model.UpdateFromFoundWorkItemsControl();
        }

        private void UnselecteAll_Click(object sender, RoutedEventArgs e)
        {
            SelectAllFoundedWorkItems(false);
            Model.UpdateFromFoundWorkItemsControl();
        }

        private void Load()
        {
            Model.UseLinkedWorkItems = Model.QueryConfiguration.UseLinkedWorkItems;
            Dispatcher.Invoke(new Action(() => SelectQuery(Model.SelectedQuery)));
        }

        private void FoundWorkItemCheckBox_CheckChanged(object sender, RoutedEventArgs e)
        {
            if (_longOperationTimer == null)
            {
                _longOperationTimer = new DispatcherTimer();
                _longOperationTimer.Interval = TimeSpan.FromMilliseconds(200);
                _longOperationTimer.Start();
                _longOperationTimer.Tick += delegate
                {
                    _longOperationTimer.Stop();
                    Model.UpdateFromFoundWorkItemsControl();

                };
            }
            else
            {
                _longOperationTimer.Stop();
                _longOperationTimer.Start();
            }
        }

        /// <summary>
        /// Select or deselect all items in founded work items check boxes in list box.
        /// </summary>
        private void SelectAllFoundedWorkItems(bool select)
        {
            foreach (DataItemModel<SynchronizedWorkItemViewModel> item in FoundList.Items)
            {
                item.IsChecked = select;
            }
        }

        /// <summary>
        /// Show or hide query tree popup.
        /// </summary>
        private void QueryBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            TreeViewPopup.IsOpen = !TreeViewPopup.IsOpen;
        }

        /// <summary>
        /// Select a query and load it if it is different to the current shown query
        /// </summary>
        private void QueryTree_OnSelected(object sender, RoutedEventArgs e)
        {
            var query = ((TreeViewItem)e.OriginalSource).DataContext as QueryItem;
            if (query != null)
            {
                if (Model.SelectedQuery != null && Model.SelectedQuery.Id == query.Id)
                {
                    return;
                }
            }

            SelectQuery(query);
        }

        private void SelectQuery(object query)
        {
            var queryItem = query as QueryItem;
            var queryFolder = query as QueryFolder;

            if (queryFolder != null || query == null)
            {
                return;
            }

            TreeViewPopup.IsOpen = false;
            QueryBox.Items.Clear();

            Model.SelectedQuery = queryItem;
            QueryBox.Items.Add(Model.SelectedQuery.Name);

            QueryBox.SelectedItem = QueryBox.Items[0];
        }

        private void ClosePopupButton_OnClick(object sender, RoutedEventArgs e)
        {
            TreeViewPopup.IsOpen = false;
        }

        private void QueryTree_Loaded(object sender, RoutedEventArgs e)
        {
            if (ActualHeight > 0)
            {
                TreeViewPopup.HorizontalOffset = 0;
                TreeViewPopup.VerticalOffset = 0;
                TreeViewPopup.MaxHeight = ActualHeight * 0.9;
            }
        }

        private void FindByTitle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return && !string.IsNullOrWhiteSpace(TitleTextBox.Text))
            {
                Find_Click(sender, e);
            }
        }
    }
}
