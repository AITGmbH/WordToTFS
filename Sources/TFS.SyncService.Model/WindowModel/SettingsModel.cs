#region Usings
using System;
using System.IO;
using AIT.TFS.SyncService.Factory;
using AIT.TFS.SyncService.Model.WindowModelBase;
using Microsoft.Win32;
using System.Diagnostics;
using System.Collections.Generic;
using AIT.TFS.SyncService.Model.Properties;
#endregion
namespace AIT.TFS.SyncService.Model.WindowModel
{

    /// <summary>
    /// The attached model to the settings dialog
    /// </summary>
    public class SettingsModel : ExtBaseModel
    {
        #region Consts
        const string WordToTFSVariableName = "WORDTOTFS";
        const string WordToTFSExecutableName = "WordToTFS.exe";
        #endregion
        #region Private fields
        private string _path;
        private string _clickOnceLocation;
        public  static TraceLevel Level;
        #endregion
        #region Public properties
        /// <summary>
        /// Gets whether debugging is turned on or off
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Binding against property.")]
        public bool IsConsoleExtensionActivated
        {
            //Check if the exe is contained in the path variable
            get
            {
                return PathContainsWordToTFSVariableName();
            }
            //Save the path of the vsto to the path 
            set
            {
                if (value)
                {
                    ConsoleExtensionIsActivated(this, null);
                    //InitializeEnvironmentVariables();
                }
                else
                {
                    RemoveEnvironmentVariable();
                }
        
            }
        }

        /// <summary>
        /// Gets list of available trace levels for logging
        /// </summary>
        public IEnumerable<TraceLevel> AvailableLevels
        {
            get
            {
                var values = Enum.GetValues(typeof(TraceLevel));
                var levels = new List<TraceLevel>();
                foreach(var level in values)
                {
                    levels.Add((TraceLevel) level);
                }
                return levels;
            }
        }

        /// <summary>
        /// Selected trace level for logging
        /// </summary>
        public TraceLevel SelectedLevel
        {
            get
            {
                return Level;
            }

            set
            {
                Level = value;
                SyncServiceTrace.DebugLevel.Level = value;
                TriggerPropertyChanged(nameof(SelectedLevel));
                if (SyncServiceTrace.DebugLevel.Level != TraceLevel.Off)
                {
                    DebugLogMessageNeeded(this, null);
                }
            }
        }

        #endregion
        #region Private methods
        /// <summary>
        /// Remov the path to the clickonce location from the path
        /// </summary>
        private void RemoveEnvironmentVariable()
        {
           
            GetPathAndClickOnceLocations();
            var foundVariableIndex = _path.IndexOf("%"+WordToTFSVariableName+"%", StringComparison.Ordinal);
            //Remove all references that point to the variable 
            if (foundVariableIndex != -1)
            {
                _path = _path.Remove(foundVariableIndex, WordToTFSVariableName.Length+2);  
            }
            //Remove all references that point directly to the path
            var currentPathFromVariable = Environment.GetEnvironmentVariable(WordToTFSVariableName, EnvironmentVariableTarget.User);
            if (!String.IsNullOrEmpty(currentPathFromVariable))
            {
                var foundPathIndex = _path.IndexOf(currentPathFromVariable, StringComparison.Ordinal);
                if (foundPathIndex != -1)
                 {
                     _path = _path.Remove(foundPathIndex, currentPathFromVariable.Length + 1);
                 }
            }
            //Remove all references that point directly to the current assembly location
            var foundCurrentAssembly = _path.IndexOf(_clickOnceLocation, StringComparison.Ordinal);
            //Remove all references that point to the variable 
            if (foundCurrentAssembly != -1)
            {
                _path = _path.Remove(foundCurrentAssembly, _clickOnceLocation.Length + 1);
            }
            Environment.SetEnvironmentVariable(WordToTFSVariableName, "", EnvironmentVariableTarget.User);
            var environment = Registry.CurrentUser.CreateSubKey(@"Environment\");
            if (environment != null)
            {
                environment.SetValue("PATH", _path, RegistryValueKind.ExpandString);
            }
        }
        /// <summary>
        /// Check if a given string contains the wordtotfsvariable
        /// </summary>
        /// <returns></returns>
        private bool PathContainsWordToTFSVariableName()
        {
            GetPathAndClickOnceLocations();

            SyncServiceTrace.I(Resources.CheckIfConsoleIsActivated);
            SyncServiceTrace.I(Resources.Location, _clickOnceLocation);
            SyncServiceTrace.I(Resources.VariableName, WordToTFSVariableName);
            SyncServiceTrace.I(Resources.Path, _path);

            if (_path == null)
            {
                SyncServiceTrace.I(Resources.ConsoleIsNotActivated);
                return false;
            }
            //Path is already set
            if (_path.Contains(_clickOnceLocation) || _path.Contains(WordToTFSVariableName))
            {
                SyncServiceTrace.I(Resources.ConsoleIsActivated);
                return true;
            }
            else
            {
                SyncServiceTrace.I(Resources.ConsoleIsNotActivated);
                return false;
            } 
        }

        /// <summary>
        /// Get the click once location and the path of the user
        /// </summary>
        private void GetPathAndClickOnceLocations()
        {
            SyncServiceTrace.I(Resources.GetPath);
            //Get the assembly information and the codebase
            var assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
            var uriCodeBase = new Uri(assemblyInfo.CodeBase);

            _clickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath);
      
            var environment = Registry.CurrentUser.CreateSubKey(@"Environment\");
            if (environment != null)
            {
                _path = (string)environment.GetValue("PATH");
            }
        }
        #endregion
        #region Public methods
        /// <summary>
        /// This method will add the path to the WordToTFS dlls and exe to the Path variable of the user contect. This make WordToTFS.exe usable from the command line
        /// </summary>
        // ReSharper disable once UnusedMethodReturnValue.Local
        public bool InitializeEnvironmentVariables()
        {

            GetPathAndClickOnceLocations();
            if (String.IsNullOrEmpty(_clickOnceLocation))
            {
                SyncServiceTrace.E(Resources.PathCanNotBeDetermined, _clickOnceLocation);
                return false;
            }

            if (!File.Exists(Path.Combine(_clickOnceLocation, WordToTFSExecutableName)))
            {
                SyncServiceTrace.E(Resources.WordToTFSNotFound, _clickOnceLocation);
                return false;
            }

            //Add the WordToTFS variable and add it to the path
            Environment.SetEnvironmentVariable(WordToTFSVariableName, _clickOnceLocation, EnvironmentVariableTarget.User);

            SyncServiceTrace.I(Resources.SetEnvironmentVariable, WordToTFSVariableName, _clickOnceLocation);

            if (PathContainsWordToTFSVariableName()) return true;
            _path += "%" + WordToTFSVariableName + "%;";
            var environment = Registry.CurrentUser.CreateSubKey(@"Environment\");
            if (environment != null)
            {
                environment.SetValue("PATH", _path, RegistryValueKind.ExpandString);
            }
            //Environment.SetEnvironmentVariable("Path", _path, EnvironmentVariableTarget);
            SyncServiceTrace.I(Resources.ExtendedPath, WordToTFSVariableName);
            SyncServiceTrace.I(Resources.NewPath, _path);

            return true;
        }
        #endregion
        #region Event handlers
        /// <summary>
        /// Occurs when the debug log is about to being activated.
        /// </summary>
        public event EventHandler DebugLogMessageNeeded;
        /// <summary>
        /// Handler to show a custom message when the console extension is activated
        /// </summary>
        public event EventHandler ConsoleExtensionIsActivated;
        #endregion
    }
}
