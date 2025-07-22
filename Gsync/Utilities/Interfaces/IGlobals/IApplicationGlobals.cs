using Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public interface IApplicationGlobals
    {
        IFileSystemFolderPaths FS { get; }
        Application OutlookApplication { get; }       
        Task<AppGlobals> InitAsync();
    }
}