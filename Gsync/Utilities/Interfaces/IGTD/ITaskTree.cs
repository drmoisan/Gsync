using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public interface ITaskTree: INotifyPropertyChanged
    {
        bool ActiveBranch { get; set; }
        string ExpandChildren { get; set; }
        string ExpandChildrenState { get; set; }
        int VisibleTreeState { get; set; }
    }
}
