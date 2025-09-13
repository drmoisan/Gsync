using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public interface ITimerFactory
    {
        ITimerWrapper CreateTimer(TimeSpan interval);
    }
}
