#region Usings
using System.Collections.Generic;
using System.Globalization;
#endregion

namespace AIT.TFS.SyncService.Model.Helper
{
    /// <summary>
    /// Class represents one element in Area path / Iteration path.
    /// </summary>
    internal class PathElement
    {
        #region Internal consts

        internal const string ConstDelimiter = "\\";
        
        #endregion Internal consts

        #region Private fields

        private List<PathElement> _childs = new List<PathElement>();
        
        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PathElement"/> class.
        /// </summary>
        /// <param name="parent">Parent object.</param>
        /// <param name="pathPart">Path part.</param>
        public PathElement(PathElement parent, string pathPart)
        {
            Parent = parent;
            PathPart = pathPart;
            Level = 1;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PathElement"/> class.
        /// </summary>
        /// <param name="parent">Parent object.</param>
        /// <param name="pathPart">Path part.</param>
        /// <param name="level">Level of the element.</param>
        private PathElement(PathElement parent, string pathPart, int level)
        {
            Parent = parent;
            PathPart = pathPart;
            Level = level;
        }

        #endregion Constructors

        #region Public properties
        /// <summary>
        /// Gets the child objects.
        /// </summary>
        public IEnumerable<PathElement> Childs
        {
            get { return _childs; }
        }

        /// <summary>
        /// Gets the level of the element. Start level is 1.
        /// </summary>
        public int Level { get; private set; }

        /// <summary>
        /// Gets the parent object.
        /// </summary>
        public PathElement Parent { get; private set; }

        /// <summary>
        /// Gets the path part.
        /// </summary>
        public string PathPart { get; private set; }

        /// <summary>
        /// The m
        /// </summary>
        public string WholePath
        {
            get
            {
                if (Parent != null)
                    return string.Join(ConstDelimiter, Parent.WholePath, PathPart);
                return PathPart;
            }
        }

        #endregion Public properties

        #region Public methods

        /// <summary>
        /// Splits the path and add first part to the childs.
        /// </summary>
        /// <param name="path">Path to split and add</param>
        public void Add(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }
            var parts = new List<string>(path.Split(ConstDelimiter[0]));
            var firstPart = parts[0];
            parts.RemoveAt(0);
            var newParts = string.Empty;
            if (parts.Count > 0)
            {
                newParts = string.Join(ConstDelimiter, parts.ToArray());
            }
            foreach (var child in _childs)
            {
                if (child.PathPart == firstPart)
                {
                    if (string.IsNullOrEmpty(newParts))
                    {
                        return;
                    }
                    child.Add(newParts);
                    return;
                }
            }
            var pathElement = new PathElement(this, firstPart, Level + 1);
            pathElement.Add(newParts);
            _childs.Add(pathElement);
            _childs.Sort((a, b) =>
                string.Compare(a.PathPart, b.PathPart, CultureInfo.CurrentCulture, CompareOptions.OrdinalIgnoreCase));
        }

        #endregion Public methods
    }
}
