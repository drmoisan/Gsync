using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public interface IProjectElement:INotifyPropertyChanged, IEquatable<IProjectElement>, IComparable<IProjectElement>
    {
        /// <summary>
        /// Unique identifier for the element.
        /// </summary>
        string ID { get; set; }
        /// <summary>
        /// Name of the element.
        /// </summary>
        string Name { get; set; }
    }    
}
