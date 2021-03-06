﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace TFS.SyncService.W2TConsole.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("TFS.SyncService.W2TConsole.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to If a configuration file is specified, it is not allowed to specify basic settings by command line parameters. You can use either the command line parameters or the configuration file for the specification of basic settings..
        /// </summary>
        internal static string CommandArgumentsError {
            get {
                return ResourceManager.GetString("CommandArgumentsError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The item {0} is missing in your configuration file..
        /// </summary>
        internal static string ConfigurationFileMisconfigured {
            get {
                return ResourceManager.GetString("ConfigurationFileMisconfigured", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The specified configuration file does not exist..
        /// </summary>
        internal static string ConfigurationFileNotExistsError {
            get {
                return ResourceManager.GetString("ConfigurationFileNotExistsError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The Word template: {0} does not exist..
        /// </summary>
        internal static string Error_TemplateDoesNotExist {
            get {
                return ResourceManager.GetString("Error_TemplateDoesNotExist", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Press any key to exit the console..
        /// </summary>
        internal static string ExitInformation {
            get {
                return ResourceManager.GetString("ExitInformation", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The specified file where to store the report already exists. Either set the overwrite option to true or specify another file name..
        /// </summary>
        internal static string FileExistsError {
            get {
                return ResourceManager.GetString("FileExistsError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The file path where to store the document was not specified..
        /// </summary>
        internal static string FilenameNotSpecifiedError {
            get {
                return ResourceManager.GetString("FilenameNotSpecifiedError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The document has been saved to {0}.
        /// </summary>
        internal static string FileSavedInformation {
            get {
                return ResourceManager.GetString("FileSavedInformation", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to There were no work items found for your arguments..
        /// </summary>
        internal static string NoWorkItemsFoundInformation {
            get {
                return ResourceManager.GetString("NoWorkItemsFoundInformation", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to There was no option specified how to get work items..
        /// </summary>
        internal static string OptionsNullError {
            get {
                return ResourceManager.GetString("OptionsNullError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Please specify either the option &quot;WorkItemIDs&quot; or the option WorkitemQuery..
        /// </summary>
        internal static string OptionsNullErrorInstruction {
            get {
                return ResourceManager.GetString("OptionsNullErrorInstruction", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to It was specified to get work items by id and by query. Only one of these two options can be chosen..
        /// </summary>
        internal static string OptionsOverloadedError {
            get {
                return ResourceManager.GetString("OptionsOverloadedError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The project was not specified..
        /// </summary>
        internal static string ProjectNotSpecifiedError {
            get {
                return ResourceManager.GetString("ProjectNotSpecifiedError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Configuration file is about being read in..
        /// </summary>
        internal static string ReadConfigurationInformation {
            get {
                return ResourceManager.GetString("ReadConfigurationInformation", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The creation of the report failed..
        /// </summary>
        internal static string ReportCreationFailed {
            get {
                return ResourceManager.GetString("ReportCreationFailed", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The report was created..
        /// </summary>
        internal static string ReportCreationSuccessfull {
            get {
                return ResourceManager.GetString("ReportCreationSuccessfull", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The server was not specified..
        /// </summary>
        internal static string ServerNotSpecifiedError {
            get {
                return ResourceManager.GetString("ServerNotSpecifiedError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The WordToTFS-template which should be used was not specified..
        /// </summary>
        internal static string TemplateNotSpecifiedError {
            get {
                return ResourceManager.GetString("TemplateNotSpecifiedError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to There was no option specified where to place the test configuration information. The available options are &quot;AboveTestPlan&quot;, &quot;BeneathTestPlan&quot;, &quot;BeneathTestSuites&quot; or &quot;BeneathFirstTestSuite&quot;..
        /// </summary>
        internal static string TestSettingsConfigurationPositionError {
            get {
                return ResourceManager.GetString("TestSettingsConfigurationPositionError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to There was no option specified how to create the document structure. The available options are &quot;IterationPath&quot;, &quot;AreaPath&quot; or &quot;TestPlanHierarchy&quot;..
        /// </summary>
        internal static string TestSettingsDocumentStructureError {
            get {
                return ResourceManager.GetString("TestSettingsDocumentStructureError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Please adapt the configuration file properly..
        /// </summary>
        internal static string TestSettingsErrorInstruction {
            get {
                return ResourceManager.GetString("TestSettingsErrorInstruction", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The option specified where to place the test configuration information is not valid. The available options are &quot;AboveTestPlan&quot;, &quot;BeneathTestPlan&quot;, &quot;BeneathTestSuites&quot; or &quot;BeneathFirstTestSuite&quot;..
        /// </summary>
        internal static string TestSettingsInvalidConfigurationPositionError {
            get {
                return ResourceManager.GetString("TestSettingsInvalidConfigurationPositionError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The option specified how to create the document structure is not valid. The available options are &quot;IterationPath&quot;, &quot;AreaPath&quot; or &quot;TestPlanHierarchy&quot;..
        /// </summary>
        internal static string TestSettingsInvalidDocumentStructureError {
            get {
                return ResourceManager.GetString("TestSettingsInvalidDocumentStructureError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The option specified how to sort the cases is not valid. cases. The available options are &quot;IterationPath&quot;, &quot;AreaPath&quot; or &quot;WorkItemId&quot;..
        /// </summary>
        internal static string TestSettingsInvalidTestCasesSortError {
            get {
                return ResourceManager.GetString("TestSettingsInvalidTestCasesSortError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The specified test plan does not exist..
        /// </summary>
        internal static string TestSettingsInvalidTestPlan {
            get {
                return ResourceManager.GetString("TestSettingsInvalidTestPlan", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The specified test suite does not exist..
        /// </summary>
        internal static string TestSettingsInvalidTestSuite {
            get {
                return ResourceManager.GetString("TestSettingsInvalidTestSuite", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to There was no option specified how to sort the test cases. The available options are &quot;IterationPath&quot;, &quot;AreaPath&quot; or &quot;WorkItemId&quot;..
        /// </summary>
        internal static string TestSettingsTestCasesSortError {
            get {
                return ResourceManager.GetString("TestSettingsTestCasesSortError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to There was no test plan specified in the configuration file..
        /// </summary>
        internal static string TestSettingsTestPlanError {
            get {
                return ResourceManager.GetString("TestSettingsTestPlanError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The document and Word have been closed.
        /// </summary>
        internal static string WordClosed {
            get {
                return ResourceManager.GetString("WordClosed", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Number of found work items:.
        /// </summary>
        internal static string WorkItemsCountInformation {
            get {
                return ResourceManager.GetString("WorkItemsCountInformation", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to All work items were imported to the document..
        /// </summary>
        internal static string WorkItemsImportedInformation {
            get {
                return ResourceManager.GetString("WorkItemsImportedInformation", resourceCulture);
            }
        }
    }
}
