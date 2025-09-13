using System;

namespace Gsync.Utilities.Interfaces
{
    public interface ITimerWrapper: IGenericTimer
    {
        bool AutoReset { get; set; }
        void ResetTimer();


    }
}