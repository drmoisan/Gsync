using System;

namespace Gsync.Utilities.Interfaces.IHelperClasses
{
    public interface IDirectory
    {
        bool Exists(string path);
        void CreateDirectory(string path);
    }
}