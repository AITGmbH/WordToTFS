#region Usings
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
#endregion

namespace AIT.TFS.SyncService.View.Controls.Helper
{
    /// <summary>
    /// Helper class with that provides methods
    /// </summary>
    public static class TreeViewHelper
    {
        private static readonly Dictionary<DependencyObject, TreeViewSelectedItemBehavior> Behaviors = new Dictionary<DependencyObject, TreeViewSelectedItemBehavior>();

        /// <summary>
        /// Gets the selected item property of a dependency object.
        /// </summary>
        /// <param name="dependencyObject">The dependency object from which to get the selected item.</param>
        /// <returns>The selected item.</returns>
        public static object GetSelectedItem(DependencyObject dependencyObject)
        {
            if (dependencyObject == null) return null;
            return dependencyObject.GetValue(SelectedItemProperty);
        }

        /// <summary>
        /// Sets the selected item property of a dependency object.
        /// </summary>
        /// <param name="dependencyObject">The dependency object.</param>
        /// <param name="value">The new value.</param>
        public static void SetSelectedItem(DependencyObject dependencyObject, object value)
        {
            if (dependencyObject != null) dependencyObject.SetValue(SelectedItemProperty, value);
        }

        /// <summary>
        /// Using a DependencyProperty as the backing store for SelectedItem.  This enables animation, styling, binding, etc...
        /// </summary>
        public static readonly DependencyProperty SelectedItemProperty = DependencyProperty.RegisterAttached("SelectedItem",
                                                                                                             typeof(object),
                                                                                                             typeof(TreeViewHelper),
                                                                                                             new UIPropertyMetadata(null, SelectedItemChanged));

        private static void SelectedItemChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            if (!(obj is TreeView)) return;

            if (!Behaviors.ContainsKey(obj)) // Not used yet
                Behaviors.Add(obj, new TreeViewSelectedItemBehavior(obj as TreeView));

            var itemBehavior = Behaviors[obj];
            itemBehavior.ChangeSelectedItem(e.NewValue);
        }

        /// <summary>
        /// Private helper class.
        /// </summary>
        private class TreeViewSelectedItemBehavior
        {
            private readonly TreeView _view;

            /// <summary>
            /// Initializes a new instance of the <see cref="TreeViewSelectedItemBehavior"/> class.
            /// </summary>
            /// <param name="view">Associated view.</param>
            public TreeViewSelectedItemBehavior(TreeView view)
            {
                _view = view;
                view.SelectedItemChanged += (sender, e) => SetSelectedItem(view, e.NewValue);
            }

            /// <summary>
            /// Selects the tree view item in tree view.
            /// </summary>
            /// <param name="itemToSelect">Item to select.</param>
            internal void ChangeSelectedItem(object itemToSelect)
            {
                if (itemToSelect == null) return;
                var item = _view.ItemContainerGenerator.ContainerFromItem(itemToSelect) as TreeViewItem;

                // if (_view.Items.Contains(p))
                if (item == null)
                {
                    for (var index = 0; index < _view.Items.Count; index++)
                    {
                        var childItem = _view.ItemContainerGenerator.ContainerFromItem(_view.Items[index]) as TreeViewItem;
                        if (childItem == null)
                        {
                            continue;
                        }

                        if (SelectInChilds(childItem, itemToSelect))
                        {
                            break;
                        }
                    }
                }
                else
                {
                    item.IsSelected = true;
                }
            }

            /// <summary>
            /// Selects the tree view item in tree view item.
            /// </summary>
            /// <param name="p">Item to select.</param>
            private static bool SelectInChilds(TreeViewItem treeViewItem, object p)
            {
                var item = treeViewItem.ItemContainerGenerator.ContainerFromItem(p) as TreeViewItem;
                if (item == null)
                {
                    for (var index = 0; index < treeViewItem.Items.Count; index++)
                    {
                        var childItem = treeViewItem.ItemContainerGenerator.ContainerFromItem(treeViewItem.Items[index]) as TreeViewItem;
                        if (childItem == null)
                        {
                            continue;
                        }

                        if (SelectInChilds(childItem, p))
                        {
                            return true;
                        }
                    }
                }
                else
                {
                    item.IsSelected = true;
                    return true;
                }

                return false;
            }
        }
    }
}