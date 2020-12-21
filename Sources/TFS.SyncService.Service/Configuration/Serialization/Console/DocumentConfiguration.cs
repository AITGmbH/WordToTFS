#region Usings
using System;
using System.IO;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Service.Properties;
using System.Diagnostics;
#endregion

namespace AIT.TFS.SyncService.Service.Configuration.Serialization.Console
{
    /// <summary>
    /// The configuration settings for the console
    /// </summary>
    [XmlRoot("DocumentConfiguration", Namespace = "", IsNullable = false)]
    public class DocumentConfiguration
    {
        #region Fields
        public Settings Settings;
        public TestSpecSettings TestSpecSettings;
        public TestResultSettings TestResultSettings;
        private Configuration _configuration;
        #endregion

        #region Constructor

        /// <summary>
        /// Initialize a configuration with given parameters.
        /// </summary>
        /// <param name="server"></param>
        /// <param name="project"></param>
        /// <param name="filename"></param>
        /// <param name="templateName"></param>
        /// <param name="overwrite"></param>
        /// <param name="closeOnFinish"></param>
        /// <param name="configFile"></param>
        /// <param name="wordHidden"></param>
        /// <param name="wordTemplateFile"></param>
        public DocumentConfiguration(string server, string project, string filename, string templateName, bool overwrite, bool closeOnFinish, string configFile, bool wordHidden, string wordTemplateFile, TraceLevel debugLevel)
        {
            Settings = new Settings();
            // Check that either XML-file or the settings-arguments are specified
            if ((server != null || project != null || filename != null || templateName != null || overwrite || closeOnFinish ) && configFile != null)
            {
                throw new ArgumentException(Resources.CommandArgumentsError);
            }

            // Read settings from XML-file
            if (configFile != null)
            {
                try
                {
                    ReadConfigurationFile(configFile);

                    // Check for the basic settings and add them
                    if (_configuration.Settings == null)
                    {
                        throw new ArgumentException(string.Format(Resources.ConfigurationFileMisconfigured, "Settings"));
                    }

                    Settings = _configuration.Settings;


                    // Check for the TestSpec Settings and add them if they exist
                    if (_configuration.TestConfiguration != null && _configuration.TestConfiguration.TestSpecificationConfiguration != null)
                    {
                        TestSpecSettings = _configuration.TestConfiguration.TestSpecificationConfiguration;

                        if (TestSpecSettings.TestSuite == null)
                        {
                            TestSpecSettings.TestSuite = TestSpecSettings.TestPlan;
                        }

                        TestResultSettings = _configuration.TestConfiguration.TestResultConfiguration;
                    }

                    // Check for the TestResultSettings and add them if they exist
                    if (_configuration.TestConfiguration != null && _configuration.TestConfiguration.TestResultConfiguration != null)
                    {
                        // Default-Werte für nicht konfigurierte (optionale) Felder setzen
                        if (TestResultSettings.TestSuite == null)
                        {
                            TestResultSettings.TestSuite = TestResultSettings.TestPlan;
                        }

                        if (TestResultSettings.Build == null)
                        {
                            TestResultSettings.Build = "All";
                        }

                        if (TestResultSettings.TestConfiguration == null)
                        {
                            TestResultSettings.TestConfiguration = "All";
                        }
                    }
                }
                catch (ArgumentException e)
                {
                    // ReSharper disable once PossibleIntendedRethrow
                    throw e;
                }

            }
            // Read settings from commandline
            else
            {
                try
                {
                    Settings = ReadConsoleArguments(server, project, filename, templateName, overwrite, closeOnFinish, wordHidden, wordTemplateFile, debugLevel);
                }
                catch (Exception e)
                {
                    // ReSharper disable once PossibleIntendedRethrow
                    throw e;
                }
            }
        }
        #endregion

        #region Private methods

        /// <summary>
        /// Reads the settings from the commandline arguments
        /// </summary>
        /// <param name="server"></param>
        /// <param name="project"></param>
        /// <param name="filename"></param>
        /// <param name="templateName"></param>
        /// <param name="overwrite"></param>
        /// <param name="closeOnFinish"></param>
        /// <param name="wordHidden">Specifies is Word should be hidden</param>
        /// <param name="wordTemplateName">The file path to a word template (.dotx)</param>
        /// Servername + collection
        /// Projectname
        /// file path, where to save the created document
        /// Template, which is to be used
        /// Specifies whether the file is overwritten if it exists or not
        /// Specifies whether the document should be closed after creation
        /// <returns></returns> Settings
        private Settings ReadConsoleArguments(string server, string project, string filename, string templateName, bool overwrite, bool closeOnFinish, bool wordHidden, string wordTemplateName, TraceLevel debugLevel)
        {
            // Check if all required arguments are given
            try
            {
                ValidateArguments(server, project, filename, templateName);
            }
            catch (ArgumentException e)
            {
                // ReSharper disable once PossibleIntendedRethrow
                throw e;
            }

            Settings.Server = server;
            Settings.Project = project;
            Settings.Filename = filename;
            Settings.Template = templateName;
            Settings.Overwrite = overwrite;
            Settings.CloseOnFinish = closeOnFinish;
            Settings.WordHidden = wordHidden;
            Settings.DotxTemplate = wordTemplateName;
            Settings.DebugLevel = (int) debugLevel;

            return Settings;
        }


        /// <summary>
        /// Reads the settings from the XML-ConfigFile
        /// </summary>
        /// <param name="configFile"></param>Path, to the XML-file
        /// <returns></returns> Settings
        private void ReadConfigurationFile(string configFile)
        {
            if (!File.Exists(configFile))
            {
                throw new ArgumentException(Resources.ConfigurationFileNotExistsError);
            }

            // Deserialize the xml file
            var serializer = new XmlSerializer(typeof(Configuration));
            using (var stream = new StreamReader(configFile))
            {
                _configuration = serializer.Deserialize(stream) as Configuration;
            }

            if (_configuration != null)
            {
                var settings = _configuration.Settings;
                // Check if all required arguments are specified in the XML-file
                try
                {
                    ValidateArguments(settings.Server, settings.Project, settings.Filename, settings.Template);
                }
                catch (Exception e)
                {
                    // ReSharper disable once PossibleIntendedRethrow
                    throw e;
                }
            }
        }


        /// <summary>
        /// Checks, whether all required arguments are specified
        /// </summary>
        /// <param name="server"></param> Servername + Collection
        /// <param name="project"></param> Projectname
        /// <param name="filename"></param> file paht, where to save the created document
        /// <param name="templateName"></param> Template, which is to be used
        private void ValidateArguments(string server, string project, string filename, string templateName)
        {
            if (server == null)
            {
                throw new ArgumentException(Resources.ServerNotSpecifiedError);
            }

            if (project == null)
            {
                throw new ArgumentException(Resources.ProjectNotSpecifiedError);
            }

            if (filename == null)
            {
                throw new ArgumentException(Resources.FilenameNotSpecifiedError);
            }

            if (templateName == null)
            {
                throw new ArgumentException(Resources.TemplateNotSpecifiedError);
            }
        }
        #endregion
    }
}
