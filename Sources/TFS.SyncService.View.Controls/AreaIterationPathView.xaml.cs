using AIT.TFS.SyncService.Contracts.WorkItemObjects;
using AIT.TFS.SyncService.Model.WindowModel;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using Microsoft.Office.Interop.Word;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for AreaIterationPathView.
    /// </summary>
    public partial class AreaIterationPathView : IHostPaneControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AreaIterationPathView"/> class.
        /// </summary>
        public AreaIterationPathView()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Gets or sets the ViewModel for this View.
        /// </summary>
        public AreaIterationPathViewModel Model
        {
            get
            {
                return DataContext as AreaIterationPathViewModel;
            }
            set
            {
                DataContext = value;
            }
        }

        /// <summary>
        /// Handles the tree view selected item changed event.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The event arguments.</param>
        private void HandleTreeViewSelectedItemChangedEvent(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            Model.SelectedAreaIterationNode = e.NewValue as IAreaIterationNode;
        }

        /// <summary>
        /// Select the last selected item.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void HandleUserControlLoadedEvent(object sender, RoutedEventArgs e)
        {
            if (AreaIterationTree.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
            {
                AreaIterationTree.ItemContainerGenerator.StatusChanged +=
                    delegate { SelectTreeAreaIterationPath(AreaIterationTree, Model.PersistedAreaIterationPath); };
            }
            else
            {
                SelectTreeAreaIterationPath(AreaIterationTree, Model.PersistedAreaIterationPath);
            }
        }

        /// <summary>
        /// Handles the insert click event.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void HandleInsertClickEvent(object sender, RoutedEventArgs e)
        {
            Model.DoInsert();
        }

        /// <summary>
        /// Selects the node with the given path in the tree view.
        /// </summary>
        /// <param name="parentContainer">The parent container.</param>
        /// <param name="path">The path of the node to select.</param>
        private static void SelectTreeAreaIterationPath(ItemsControl parentContainer, string path)
        {
            if (parentContainer == null)
            {
                return;
            }

            var areaIterationNode = parentContainer.DataContext as IAreaIterationNode;
            if (areaIterationNode != null && areaIterationNode.Path == path)
            {
                var treeViewItem = parentContainer as TreeViewItem;
                if (treeViewItem != null)
                {
                    treeViewItem.BringIntoView();
                    treeViewItem.IsSelected = true;
                    return;
                }
            }

            foreach (object item in parentContainer.Items)
            {
                var currentContainer =
                    parentContainer.ItemContainerGenerator.ContainerFromItem(item) as TreeViewItem;
                if (currentContainer != null)
                {
                    currentContainer.IsExpanded = true;
                    if (currentContainer.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
                    {
                        currentContainer.ItemContainerGenerator.StatusChanged +=
                            delegate { SelectTreeAreaIterationPath(currentContainer, path); };
                    }
                    else
                    {
                        SelectTreeAreaIterationPath(currentContainer, path);
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets Word Document instance.
        /// </summary>
        public Document AttachedDocument
        {
            get
            {
                return Model.WordDocument;
            }
        }

        /// <summary>
        /// The method is called after the visibility of the containing panel has changed.
        /// </summary>
        /// <param name="isVisible"><c>true</c>if the panel containing this view is shown. False otherwise.</param>
        public void VisibilityChanged(bool isVisible)
        {
            // We don't care
        }
    }
}
