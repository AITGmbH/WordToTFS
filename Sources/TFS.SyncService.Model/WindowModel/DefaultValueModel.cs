#region Usings
using AIT.TFS.SyncService.Model.WindowModelBase;
#endregion
namespace AIT.TFS.SyncService.Model.WindowModel
{
    /// <summary>
    /// Class encapsulates the implementation of default value model used in list view.
    /// </summary>
    public class DefaultValueModel : ExtBaseModel
    {
        #region Private fields
        private string _name;
        private string _referenceFieldName;
        private string _value;
        private string _valueVisibility;
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultValueModel" /> class.
        /// </summary>
        /// <param name="referenceFieldName">Name of the field.</param>
        /// <param name="name">Name of the default value.</param>
        /// <param name="value">Value of the default value.</param>
        /// <param name="valueVisibility">The value visibility.</param>
        /// <param name="workItemType">Type of the work item.</param>
        public DefaultValueModel(string referenceFieldName, string name, string value, string valueVisibility, string workItemType)
        {
            _referenceFieldName = referenceFieldName;
            _name = name;
            _value = value;
            _valueVisibility = valueVisibility;
            WorkItemType = workItemType;
        }
        #endregion
        #region Public properties
        /// <summary>
        /// Gets or sets the name of the field.
        /// </summary>
        public string ReferenceFieldName
        {
            get
            {
                return _referenceFieldName;
            }
            set
            {
                if (_referenceFieldName != value)
                {
                    _referenceFieldName = value;
                    TriggerPropertyChanged(this, nameof(ReferenceFieldName));
                }
            }
        }

        /// <summary>
        /// Gets or sets the name of the default value.
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    TriggerPropertyChanged(this, nameof(Name));
                }
            }
        }

        /// <summary>
        /// Gets or sets the value of the default value.
        /// </summary>
        public string Value
        {
            get { return _value; }
            set
            {
                if (_value != value)
                {
                    _value = value;
                    TriggerPropertyChanged(this, nameof(Value));
                }
            }
        }

        /// <summary>
        /// Gets or sets the visibility state of a model
        /// </summary>
        public string ValueVisibility
        {
            get
            {
                return _valueVisibility;
            }
            set
            {
                if (_valueVisibility != value)
                {
                    _valueVisibility = value;
                    TriggerPropertyChanged(this, nameof(ValueVisibility));
                }
            }
        }

        /// <summary>
        /// Gets the bold attribute for font
        /// </summary>
        public string FontWeight
        {
            get
            {
                return ValueVisibility != "Visible" ? "Bold" : "Normal";
            }
        }

        /// <summary>
        /// Gets or sets the work item type of a model
        /// </summary>
        public string WorkItemType { get; set; }
        #endregion
    }
}