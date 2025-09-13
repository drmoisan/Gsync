
using System.Collections.Concurrent;

namespace Gsync.Utilities.Interfaces
{
    public interface IFileSystemFolderPaths
    {
        ConcurrentDictionary<string, string> SpecialFolders { get; }
        void Reload();
        IAppStagingFilenames Filenames { get; }
        string MatchBestSpecialFolder(string path);
    }
}