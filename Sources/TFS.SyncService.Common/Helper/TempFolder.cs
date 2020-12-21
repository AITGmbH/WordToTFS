namespace AIT.TFS.SyncService.Common.Helper
{
    #region Usings
    using System.Collections.Generic;
    using System.IO;
    #endregion

    /// <summary>
    /// The class implements methods to work with temporary folder and file.
    /// </summary>
    public static class TempFolder
    {
        #region Private static consts
        
        private const string ConstWordToTfsSubfolder = "WordToTFS";
        
        #endregion Private static consts

        #region Private static fields

        private static Dictionary<string, string> _createdCopies = new Dictionary<string, string>();

        #endregion Private static fields

        #region Public static properties

        /// <summary>
        /// Gets main temporary folder for WordToTFS. In this folder should be created all temporary files
        /// At the start up time should be every file deleted.
        /// The folder is 'WordToTFS' in folder 'C:\Users\&lt;user&gt;\AppData\Local\Temp'.
        /// </summary>
        public static string TemporaryFolder
        {
            get
            {
                var path = Path.Combine(Path.GetTempPath(), ConstWordToTfsSubfolder);
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return path;
            }
        }
        
        #endregion Public static properties

        #region Public static methods

        /// <summary>
        /// The method gets the name of temporary file in <see cref="TemporaryFolder"/>.
        /// </summary>
        /// <returns>Name of the temporary file in <see cref="TemporaryFolder"/>.</returns>
        public static string GetTemporaryFile()
        {
            var tempFile = string.Empty;
            do
            {
                tempFile = Path.Combine(TemporaryFolder, Path.GetRandomFileName()+".tmp");
            }
            while (File.Exists(tempFile));
            return tempFile;
        }

        /// <summary>
        /// The method creates the empty temporary file in <see cref="TemporaryFolder"/>.
        /// </summary>
        /// <returns>Name of the temporary file in <see cref="TemporaryFolder"/>.</returns>
        public static string CreateTemporaryFile()
        {
            var file = GetTemporaryFile();
            File.Create(file).Close();
            return file;
        }

        /// <summary>
        /// The method creates the temporary file in <see cref="TemporaryFolder"/> with identical content as given file.
        /// </summary>
        /// <param name="originalFile">Content of this file is copied to created temporary file.</param>
        /// <returns>Name of the temporary file in <see cref="TemporaryFolder"/>.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        public static string CreateTemporaryFile(string originalFile)
        {
            if (!File.Exists(originalFile))
            {
                return CreateTemporaryFile();
            }
            if (_createdCopies.ContainsKey(originalFile))
            {
                return _createdCopies[originalFile];
            }
            var target = GetTemporaryFile();
            _createdCopies[originalFile] = target;
            try
            {
                using (var sourceStream = File.OpenRead(originalFile))
                {
                    using (var targetStream = File.OpenWrite(target))
                    {
                        // Copy the file in 1kB blocks.
                        var buffer = new byte[1024];
                        var readBytes = 0;
                        do
                        {
                            readBytes = sourceStream.Read(buffer, 0, buffer.Length);
                            if (readBytes > 0)
                            {
                                targetStream.Write(buffer, 0, readBytes);
                            }
                        }
                        while (readBytes == buffer.Length);
                    }
                }
            }
            catch
            {
                _createdCopies[originalFile] = string.Empty;
            }
            return _createdCopies[originalFile];
        }

        /// <summary>
        /// The method clear all files and subfolder in the <see cref="TemporaryFolder"/> folder.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "We need to catch all exceptions")]
        public static void ClearTempFolder()
        {
            var folder = TemporaryFolder;
            if (!Directory.Exists(folder))
            {
                return;
            }
            try
            {
                foreach (var directory in Directory.GetDirectories(folder))
                {
                    DeleteDirectory(directory);
                }  
                RemoveFilesFromFolder(folder);
            }
            catch
            {
            }
        }

        /// <summary>
        /// The method creates a subfolder in <see cref="TemporaryFolder"/> folder.
        /// </summary>
        /// <param name="prefix">Text will be used as prefix for created subfolder.</param>
        /// <returns>New subfolder in <see cref="TemporaryFolder"/> folder.</returns>
        public static string CreateNewTempFolder(string prefix)
        {
            if (prefix == null)
            {
                prefix = string.Empty;
            }
            // We will use the temp file name as part of name of created subfolder.
            var newFolder = Path.Combine(TemporaryFolder, prefix, Path.GetRandomFileName());
            while (Directory.Exists(newFolder))
            {
                newFolder = Path.Combine(TemporaryFolder, prefix, Path.GetRandomFileName());
            }
            Directory.CreateDirectory(newFolder);
            return newFolder;
        }

        #endregion Public static methods

        #region Private static methods
        
        /// <summary>
        /// The method removes every file and subfolder in given folder.
        /// </summary>
        /// <param name="folder">Folder to clear.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "We need to catch all exceptions")]
        private static void DeleteDirectory(string folder)
        {
            try
            {
                foreach (var directory in Directory.GetDirectories(folder))
                {
                    DeleteDirectory(directory);
                }
                RemoveFilesFromFolder(folder);
                Directory.Delete(folder);
            }
            catch
            {
            }
        }

        /// <summary>
        /// The method deletes all files in given folder.
        /// </summary>
        /// <param name="folder">Folder where all files should be deleted.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "We need to catch all exceptions")]
        private static void RemoveFilesFromFolder(string folder)
        {
            try
            {
                foreach (var file in Directory.GetFiles(folder))
                    File.Delete(file);
            }
            catch
            {
            }
        }

        #endregion Private static methods
    }
}
