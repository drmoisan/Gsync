using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public interface ITaskLegacy: INotifyPropertyChanged
    {
        bool Bullpin { get; set; }
        string MetaTaskLvl { get; set; }
        string MetaTaskSubject { get; set; }
        Func<string, string> ProjectsToPrograms { get; set; }
        string TaskSubject { get; set; }
        bool Today { get; set; }
        string ToDoID { get; set; }
        Task ForceSave();
        object GetItem();
        bool get_PA_FieldExists(string PA_Schema);
        bool get_VisibleTreeStateLVL(int Lvl);
        void set_VisibleTreeStateLVL(int Lvl, bool value);
        void SplitID();
        Task WriteFlagsBatch();
        void WriteFlagsBatch(Enums.GtdFlags flagsToSet);
    }
}
