#region Usings
using System;
using System.Collections.Generic;
using System.Linq;
using AIT.TFS.SyncService.Contracts.WorkItemObjects;
#endregion

namespace AIT.TFS.SyncService.Adapter.TFS2012.WorkItemObjects
{
    /// <summary>
    /// Represents one Area or Iteration node from hierarchy.
    /// </summary>
    public class AreaIterationNode : IAreaIterationNode
    {
        private readonly IList<IAreaIterationNode> _children = new List<IAreaIterationNode>();

        /// <summary>
        /// Initializes a new instance of the <see cref="AreaIterationNode"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="path">The path.</param>
        /// <param name="nodeId">The node id.</param>
        /// <param name="rootNode">The root node.</param>
        public AreaIterationNode(string name, string path, string nodeId, IAreaIterationNode rootNode)
        {
            Name = name;
            Path = path;
            NodeID = nodeId;
            RootNode = rootNode;
        }

        /// <summary>
        /// Gets name of Area/Iteration node.
        /// </summary>
        public string Name
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets path of Area/Iteration node.
        /// </summary>
        public string Path
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets path trimmed with project Area or Iteration path.
        /// </summary>
        public string RelativePath
        {
            get
            {
                if (RootNode != null && Path.StartsWith(RootNode.Path, StringComparison.OrdinalIgnoreCase))
                {
                    return Path.Substring(RootNode.Path.Length);
                }
                
                return RootNode == null ? "\\" : Path;
            }
        }

        /// <summary>
        /// Gets node ID of Area/Iteration node.
        /// </summary>
        public string NodeID
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets root node (can be Area or Iteration)
        /// </summary>
        public IAreaIterationNode RootNode
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets root node project name.
        /// </summary>
        public string RootNodeProject
        {
            get
            {
                var items = RootNode == null ? Path.Split('\\') : RootNode.Path.Split('\\');
                var projectItem = (from s in items
                                  where string.IsNullOrEmpty(s) == false
                                  select s).FirstOrDefault();
                return projectItem;
            }
        }

        /// <summary>
        /// Gets all child Area/Iteration nodes.
        /// </summary>
        public IList<IAreaIterationNode> Childs
        {
            get
            {
                return _children;
            }
        }

        /// <summary>
        /// Gets flatten hierarchy from node.
        /// </summary>
        /// <returns>A List of all descendants in the tree.</returns>
        public IEnumerable<IAreaIterationNode> FlattenHierarchy()
        {
            return FlattenHierarchy(this);
        }

        /// <summary>
        /// Flattens the hierarchy.
        /// </summary>
        /// <param name="node">The root node of the recursion.</param>
        /// <returns>A List of all descendants in the tree.</returns>
        private IEnumerable<IAreaIterationNode> FlattenHierarchy(IAreaIterationNode node)
        {
            yield return node;
            if (node.Childs != null)
            {
                foreach (var child in node.Childs)
                {
                    foreach (var childOrDescendant in child.FlattenHierarchy())
                    {
                        yield return childOrDescendant;
                    }
                }
            }
        }
    }
}