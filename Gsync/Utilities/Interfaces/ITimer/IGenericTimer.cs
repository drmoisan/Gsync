using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    /// <summary>
    /// This is a wrapper for the EventArgs class that is similar to 
    /// ElapsedEventArgs but has a public constructor.
    /// https://stackoverflow.com/questions/8940982/how-can-i-run-the-event-handler-assigned-to-a-mock
    /// </summary>
    public class TimeElapsedEventArgs : EventArgs
    {
        public DateTime SignalTime { get; private set; }

        public TimeElapsedEventArgs() : this(DateTime.Now)
        {
        }

        public TimeElapsedEventArgs(DateTime signalTime)
        {
            this.SignalTime = signalTime;
        }
    }
    
    /// <summary>
    /// This interface is a wrapper for the System.Timers.Timer class.
    /// https://stackoverflow.com/questions/8940982/how-can-i-run-the-event-handler-assigned-to-a-mock
    /// </summary>
    public interface IGenericTimer : IDisposable
    {
        double IntervalInMilliseconds { get; set; }

        TimeSpan Interval { get; set; }

        event EventHandler<TimeElapsedEventArgs> Elapsed;

        void StartTimer();

        void StopTimer();

        bool Enabled { get; set; }
    }

}
