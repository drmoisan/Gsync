using Gsync.OutlookInterop;
using Microsoft.Office.Interop.Outlook;
using System.Threading;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public interface IApplicationGlobals
    {
        IFileSystemFolderPaths FS { get; }
        Application OutlookApplication { get; }       
        Task<AppGlobals> InitAsync(SynchronizationContext context, int uiThreadId);
        public StoresWrapper Stores { get; }
    }
}