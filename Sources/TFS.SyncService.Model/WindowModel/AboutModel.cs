#region Usings
using AIT.TFS.SyncService.Common;
using AIT.TFS.SyncService.Model.WindowModelBase;
#endregion
namespace AIT.TFS.SyncService.Model.WindowModel
{
    /// <summary>
    /// Model class for about window
    /// </summary>
    public class AboutModel : ExtBaseModel
    {
        #region Public properties
        /// <summary>
        /// Gets the assembly name.
        /// </summary>
        public static string AssemblyName
        {
            get
            {
                return ProductInformation.AssemblyName;
            }
        }

        /// <summary>
        /// Gets the assembly version
        /// </summary>
        public static string AssemblyVersion
        {
            get
            {
                return ProductInformation.AssemblyVersion;
            }
        }

        /// <summary>
        /// Gets the assembly copyright.
        /// </summary>
        public static string AssemblyCopyright
        {
            get
            {
                return ProductInformation.AssemblyCopyright;
            }
        }

        /// <summary>
        /// Gets the assembly description.
        /// </summary>
        /// <value>The assembly description.</value>
        public static string AssemblyDescription
        {
            get
            {
                return ProductInformation.AssemblyDescription;
            }
        }
        #endregion
    }
}