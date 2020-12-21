namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    /// <summary>
    /// Implementation of temporary property value provider.
    /// </summary>
    internal class TfsTemporaryProvider : TfsPropertyValueProvider
    {
        private readonly object _propertyProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsTemporaryProvider"/> class.
        /// </summary>
        /// <param name="propertyProvider">Object contains the examined properties.</param>
        public TfsTemporaryProvider(object propertyProvider)
        {
            _propertyProvider = propertyProvider;
        }

        /// <summary>
        /// Gets the object which is used to determine value of property.
        /// </summary>
        public override object AssociatedObject
        {
            get { return _propertyProvider; }
        }
    }
}
