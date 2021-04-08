using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

[assembly: AssemblyDescription("Local Build")]
[assembly: CLSCompliant(false)]
[assembly: InternalsVisibleTo("TFS.SyncService.Test.Unit")]
[assembly: InternalsVisibleTo("TFS.SyncService.Test.Integration")]