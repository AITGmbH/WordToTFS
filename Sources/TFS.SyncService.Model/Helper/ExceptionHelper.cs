#region Usings
using System;
using AIT.TFS.SyncService.Contracts.InfoStorage;
using AIT.TFS.SyncService.Factory;
#endregion

namespace AIT.TFS.SyncService.Model.Helper
{
    /// <summary>
    /// Helper class that logs to file and shows notification for user
    /// </summary>
    public static class ExceptionHelper
    {
        /// <summary>
        /// Logs an exception to file and creates an entry in the info storage service.
        /// If the exception is null, nothing is logged.
        /// </summary>
        /// <param name="exception">Exception that occurred or null.</param>
        /// <param name="title">Title of the user information entry.</param>
        public static void NotifyIfException(this Exception exception, string title)
        {
            if (exception != null)
            {
                SyncServiceTrace.LogException(exception);

                var infoStorage = SyncServiceFactory.GetService<IInfoStorageService>();
                if (infoStorage != null)
                {
                    infoStorage.NotifyError(title, exception);
                }
            }
        }
    }
}