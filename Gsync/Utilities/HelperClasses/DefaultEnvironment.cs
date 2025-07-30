using System;
using Gsync.Utilities.Interfaces;

namespace Gsync.Utilities.HelperClasses
{
    public class DefaultEnvironment : IEnvironment
    {
        public Environment.SpecialFolder SpecialFolder { get; set; }

        public string GetFolderPath(Environment.SpecialFolder folder) => Environment.GetFolderPath(folder);

        public string GetEnvironmentVariable(string variable) => Environment.GetEnvironmentVariable(variable);
    }
}