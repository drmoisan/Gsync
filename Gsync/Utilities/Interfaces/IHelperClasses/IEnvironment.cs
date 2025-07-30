using System;

namespace Gsync.Utilities.Interfaces
{
    public interface IEnvironment
    {
        Environment.SpecialFolder SpecialFolder { get; set; }
        string GetFolderPath(Environment.SpecialFolder folder);
        string GetEnvironmentVariable(string variable);
    }
}