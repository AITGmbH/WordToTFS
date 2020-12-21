using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AIT.TFS.SyncService.Contracts.WorkItemObjects
{
    /// <summary>
    /// Represents one Area or Iteration node from hierarchy.
    /// </summary>
    public interface IAreaIterationNode
    {
        /// <summary>
        /// Gets name of Area/Iteration node.
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Gets path of Area/Iteration node.
        /// </summary>
        string Path { get; }
        /// <summary>
        /// Gets path trimmed with project Area or Iteration path.
        /// </summary>
        string RelativePath { get; }
        /// <summary>
        /// Gets node ID of Area/Iteration node.
        /// </summary>
        string NodeID { get; }
        /// <summary>
        /// Gets root node (can be Area or Iteration)
        /// </summary>
        IAreaIterationNode RootNode { get; }
        /// <summary>
        /// Gets root node project name.
        /// </summary>
        string RootNodeProject { get; }
        /// <summary>
        /// Gets all child Area/Iteration nodes.
        /// </summary>
        IList<IAreaIterationNode> Childs { get; }
        /// <summary>
        /// Gets flatten hierarchy from node.
        /// </summary>
        IEnumerable<IAreaIterationNode> FlattenHierarchy();
    }
}
