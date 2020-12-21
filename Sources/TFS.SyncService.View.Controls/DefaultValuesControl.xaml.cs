using AIT.TFS.SyncService.Contracts.Model;
using AIT.TFS.SyncService.Model.WindowModel;
using AIT.TFS.SyncService.View.Controls.Interfaces;
using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace AIT.TFS.SyncService.View.Controls
{
    /// <summary>
    /// Interaction logic for DefaultValuesControl.
    /// </summary>
    public partial class DefaultValuesControl : IHostPaneControl
    {
        #region Fields
        private readonly DefaultValuesModel _model;
        #endregion Fields

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultValuesControl"/> class.
        /// </summary>
        /// <param name="documentModel">Model for associated document to show the default values from.</param>
        public DefaultValuesControl(ISyncServiceDocumentModel documentModel)
        {
            if (documentModel == null)
                throw new ArgumentNullException("documentModel");
            InitializeComponent();
            WordDocument = documentModel.WordDocument as Document;
            DocumentModel = documentModel;
            _model = new DefaultValuesModel(DocumentModel);
            DataContext = _model;
        }
        #endregion Constructor

        /// <summary>
        /// Gets or sets Word Document instance.
        /// </summary>
        public Document WordDocument { get; private set; }

        /// <summary>
        /// Gets the model of associated word document.
        /// </summary>
        public ISyncServiceDocumentModel DocumentModel { get; private set; }

        #region Handled events
        
        /// <summary>
        /// Called when either the <see cref="FrameworkElement.ActualHeight"/> or the <see cref="FrameworkElement.ActualWidth"/> properties change value on this element.
        /// </summary>
        /// <param name="sender">The object where the event handler is attached.</param>
        /// <param name="e">The event data.</param>
        private void HandleControlSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!e.WidthChanged)
                return;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var view = listView.View as GridView;
                var border = VisualTreeHelper.GetChild(listView, 0) as Decorator;
                if (view == null || border == null)
                    return;
                var scroller = border.Child as ScrollViewer;
                if (scroller == null)
                    return;
                var presenter = scroller.Content as ItemsPresenter;
                if (presenter == null)
                    return;
                _model.Column1Width = (presenter.ActualWidth / 2) - 3;
                _model.Column2Width = (presenter.ActualWidth / 2) - 3;
            }));
        }

        /// <summary>
        /// Called when a button 'buttonResetValues' is clicked.
        /// </summary>
        /// <param name="sender">The object where the event handler is attached.</param>
        /// <param name="e">The event data.</param>
        private void HandleButtonResetValuesClick(object sender, RoutedEventArgs e)
        {
            _model.ResetDefaultValues();
        }
        
        /// <summary>
        /// Called when the selection of a Selector changes.
        /// </summary>
        /// <param name="sender">The object where the event handler is attached.</param>
        /// <param name="e">The event data.</param>
        private void HandleListViewSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Set the focus to the TextBox to edit the value.
            var lvi = listView.ItemContainerGenerator.ContainerFromIndex(listView.SelectedIndex) as ListViewItem;
            if (lvi == null)
                return;
            var textBoxFind = GetDescendantByType(lvi, typeof(TextBox), "textBoxValue") as TextBox;
            var uiElement = Keyboard.FocusedElement as UIElement;
            if (textBoxFind != null && textBoxFind != uiElement)
            {
                textBoxFind.Focus();
            }
        }
        
        /// <summary>
        /// Called when the keyboard is focused on this element.
        /// </summary>
        /// <param name="sender">The object where the event handler is attached.</param>
        /// <param name="e">The event data.</param>
        private void HandleControlPreviewGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            // Set the focus to the appropriate list view item.
            IInputElement element = e.NewFocus;
            var uiElement = element as UIElement;
            var tb = element as TextBox;

            if (tb == null)
                return;

            var lvi = GetParentByType(uiElement, typeof(ListViewItem)) as ListViewItem;
            if (lvi != null)
            {
                int newIndex = listView.ItemContainerGenerator.IndexFromContainer(lvi);
                listView.SelectedIndex = newIndex;
            }
        }
        
        /// <summary>
        /// Called when a key is released while focus is on this element.
        /// </summary>
        /// <param name="sender">The object where the event handler is attached.</param>
        /// <param name="e">The event data.</param>
        private void HandleControlPreviewKeyUp(object sender, KeyEventArgs e)
        {
            var textBox = Keyboard.FocusedElement as TextBox;
            if (textBox == null)
                return;

            var lvi = GetParentByType(textBox, typeof(ListViewItem)) as ListViewItem;
            int lviIndex = listView.ItemContainerGenerator.IndexFromContainer(lvi);
            if (e.Key == Key.Down || e.Key == Key.Enter || e.Key == Key.Return)
            {
                lviIndex++;
                if (lviIndex >= listView.Items.Count)
                    lviIndex = 0;
            }
            else if (e.Key == Key.Up)
            {
                lviIndex--;
                if (lviIndex < 0)
                    lviIndex = listView.Items.Count - 1;
            }

            if (listView.SelectedIndex != lviIndex)
                listView.SelectedIndex = lviIndex;
        }
        
        #endregion Handled events

        #region Private methods
 
        /// <summary>
        /// Finds the child.
        /// </summary>
        private static Visual GetDescendantByType(Visual element, Type type, string name)
        {
            if (element == null)
                return null;
            var frameworkElement = element as FrameworkElement;
            if (element.GetType() == type)
            {
                if (frameworkElement != null)
                {
                    if (frameworkElement.Name == name)
                    {
                        return frameworkElement;
                    }
                }
            }

            Visual foundElement = null;
            if (frameworkElement != null)
                frameworkElement.ApplyTemplate();

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(element); i++)
            {
                var visual = VisualTreeHelper.GetChild(element, i) as Visual;
                foundElement = GetDescendantByType(visual, type, name);
                if (foundElement != null)
                    break;
            }

            return foundElement;
        }
        
        /// <summary>
        /// Finds the parent.
        /// </summary>
        private static Visual GetParentByType(Visual element, Type type)
        {
            if (element == null)
                return null;
            if (element.GetType() == type)
            {
                return element;
            }

            var frameworkElement = element as FrameworkElement;
            if (frameworkElement != null)
                frameworkElement.ApplyTemplate();
            var visual = VisualTreeHelper.GetParent(element) as Visual;
            return GetParentByType(visual, type);
        }
 
        #endregion Private methods

        #region Implementation of IHostPaneControl
        
        /// <summary>
        /// Gets Word document attached to control.
        /// </summary>
        public Document AttachedDocument
        {
            get
            {
                return WordDocument;
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

        #endregion Implementation of IHostPaneControl
    }
}