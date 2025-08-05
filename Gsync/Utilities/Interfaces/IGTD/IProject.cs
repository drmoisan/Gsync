
using System;
using System.ComponentModel;

namespace Gsync.Utilities.Interfaces
{
    public interface IProject : IComparable<IProject>, IEquatable<IProject>, IComparable, INotifyPropertyChanged
    {        
        IProjectElement Project { get; set; }
        IProjectElement Program { get; set; }
    }
}
