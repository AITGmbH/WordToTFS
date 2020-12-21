namespace AIT.TFS.SyncService.Common
{
    #region Usings
    using System;
    using System.Deployment.Application;
    using System.Reflection;
    #endregion

    public static class ProductInformation
    {
        #region Fields

        private static Assembly _representativeAssembly;

        #endregion Fields

        #region Public properties
        /// <summary>
        /// Gets the assembly name.
        /// </summary>
        public static string AssemblyName
        {
            get
            {
                // Get all Description attributes on this assembly
                object[] attributes = RepresentativeAssembly.GetCustomAttributes(
                    typeof(AssemblyTitleAttribute), false);

                // If there aren't any Description attributes, return an empty string
                if (attributes.Length == 0)
                    return string.Empty;

                // If there is a Description attribute, return its value
                return ((AssemblyTitleAttribute)attributes[0]).Title + Environment.NewLine + "Workitem Exchange Tool";
            }
        }

        /// <summary>
        /// Gets the assembly version
        /// </summary>
        public static string AssemblyVersion
        {
            get
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    return ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                }

                // Get all Description attributes on this assembly
                object[] attributes = RepresentativeAssembly.GetCustomAttributes(
                    typeof(AssemblyVersionAttribute), false);

                // If there aren't any Description attributes, return an empty string
                if (attributes.Length == 0)
                    return RepresentativeAssembly.GetName().Version.ToString();

                // If there is a Description attribute, return its value
                return ((AssemblyVersionAttribute)attributes[0]).Version;
            }
        }

        /// <summary>
        /// Gets the assembly copyright.
        /// </summary>
        public static string AssemblyCopyright
        {
            get
            {
                // Get all Description attributes on this assembly
                object[] attributes = RepresentativeAssembly.GetCustomAttributes(
                    typeof(AssemblyCopyrightAttribute), false);

                // If there aren't any Description attributes, return an empty string
                if (attributes.Length == 0)
                    return string.Empty;

                // If there is a Description attribute, return its value
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
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
                // Get all Description attributes on this assembly
                object[] attributes = RepresentativeAssembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute),
                                                                                 false);

                // If there aren't any Description attributes, return an empty string
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }

                // If there is a Description attribute, return its value
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }


        #endregion Public properties

        #region Private properties
        /// <summary>
        /// Gets the representative assembly.
        /// </summary>
        /// <value>The representative assembly.</value>
        private static Assembly RepresentativeAssembly
        {
            get
            {
                if (_representativeAssembly == null)
                {
                    _representativeAssembly = Assembly.GetEntryAssembly();

                    if (_representativeAssembly == null)
                    {
                        _representativeAssembly = Assembly.GetExecutingAssembly();
                    }
                }

                return _representativeAssembly;
            }
        }
        #endregion Private properties
    }
}
