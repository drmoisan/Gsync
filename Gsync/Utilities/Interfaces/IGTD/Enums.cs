using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public static class Enums
    {
        public enum Priority
        {
            SomeDayMaybe = 0,
            Low = 1,
            Medium = 2,
            High = 3,
            Today = 4
        }

        [Flags]
        public enum GtdFlags
        {
            None = 0,
            Context = 1,
            People = 2,
            Projects = 4,
            Program = 8,
            Topics = 16,
            Priority = 32,
            Taskname = 64,
            Worktime = 128,
            Today = 256,
            Bullpin = 512,
            Kbf = 1024,
            DueDate = 2048,
            Reminder = 4096,
            All = 8191
        }
    }
}
