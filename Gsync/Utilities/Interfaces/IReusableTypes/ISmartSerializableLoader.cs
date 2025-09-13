using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public interface ISmartSerializableLoader<U> where U : class
    {
        ISmartSerializableConfig Config { get; }
    }

}
