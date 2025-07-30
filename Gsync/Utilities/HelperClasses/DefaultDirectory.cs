using System.IO;
using Gsync.Utilities.Interfaces.IHelperClasses;

namespace Gsync.Utilities.HelperClasses
{
    public class DefaultDirectory : IDirectory
    {
        public bool Exists(string path) => Directory.Exists(path);
        public void CreateDirectory(string path) => Directory.CreateDirectory(path);
    }
}