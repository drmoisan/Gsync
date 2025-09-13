using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public interface ITask: INotifyPropertyChanged
    {
        bool Complete { get; set; }
        DateTime CreationDate { get; set; }
        string Description { get; set; }
        DateTime DueDate { get; set; }
        object InnerObject { get; set; }
        IIDList IdList { get; set; }
        bool IdAutoCoding { get; set; }
        string InFolder { get; }
        bool ReadOnly { get; set; }
        DateTime ReminderTime { get; set; }
        DateTime StartDate { get; set; }
        int TotalWork { get; set; }

        void Synchronize();
        void Synchronize(Enums.GtdFlags flags);
    }
}
