using System;
using System.Reflection;
using System.Runtime.CompilerServices;

[assembly: AssemblyDescription("Local Build")]
[assembly: CLSCompliant(false)]
[assembly: InternalsVisibleTo("TFS.SyncService.Test.Integration")]
[assembly: InternalsVisibleTo("TFS.SyncService.Test.Unit")]