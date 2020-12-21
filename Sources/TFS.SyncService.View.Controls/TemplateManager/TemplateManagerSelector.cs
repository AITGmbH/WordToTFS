namespace AIT.TFS.SyncService.View.Controls.TemplateManager
{
    #region Usings
    using System.Windows;
    using System.Windows.Controls;
    using AIT.TFS.SyncService.Contracts.TemplateManager;
    #endregion

    /// <summary>
    /// Selector for Template Bundle and Template
    /// </summary>
    public class TemplateManagerSelector : DataTemplateSelector
    {
        /// <summary>
        /// Template for template bundle
        /// </summary>
        public DataTemplate TemplateBundlesTemplate
        {
            get;
            set;
        }

        /// <summary>
        /// Template for template.
        /// </summary>
        public DataTemplate TemplateTemplate
        {
            get;
            set;
        }

        /// <summary>
        /// Template for project mapped template.
        /// </summary>
        public DataTemplate ProjectMappedTemplateTemplate
        {
            get;
            set;
        }

        /// <summary>
        /// When overridden in a derived class, returns a System.Windows.DataTemplate based on custom logic.
        /// </summary>
        /// <param name="item">The data object for which to select the template.</param>
        /// <param name="container">The data-bound object.</param>
        /// <returns>Returns a System.Windows.DataTemplate or null. The default value is null.</returns>
        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (item is ITemplateBundle)
                return TemplateBundlesTemplate;
            if (item is ITemplate)
                return ((ITemplate) item).ProjectName == null ? TemplateTemplate : ProjectMappedTemplateTemplate;
            return base.SelectTemplate(item, container);
        }
    }
}
