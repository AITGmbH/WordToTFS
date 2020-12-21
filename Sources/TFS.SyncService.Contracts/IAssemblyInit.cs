using System;

namespace AIT.TFS.SyncService.Contracts
{
    /// <summary>
    /// Interface defines the behavior of initialization of one assembly for 'ClickOnce Deployment'.
    /// </summary>
    /// <remarks>
    /// If the initialization of assembly is required, it is to define one singleton class that derived from this interface.
    /// </remarks>
    public interface IAssemblyInit : IDisposable
    {
        /// <summary>
        /// Method initializes the assembly.
        /// </summary>
        void Init();
    }
}