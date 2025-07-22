using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Gsync.Utilities.Threading;

namespace Gsync.Utilities.HelperClasses
{
    public class TimedAsyncTask
    {
        public TimedAsyncTask(TimeSpan frequency)
        {
            _frequency = frequency;
        }

        public TimedAsyncTask(TimeSpan frequency, Func<Task> action)
        {
            _frequency = frequency;
            _action = action;
        }

        public void SetAction(Func<Task> action)
        {
            Interlocked.CompareExchange(ref _action, action, null);
        }
        private Func<Task> _action;

        private TimeSpan _frequency;
        private ThreadSafeSingleShotGuard _taskRequested = new();
        private TimerWrapper _timer;

        public void ResetTimer()
        {
            _timer?.ResetTimer();
        }

        public void CancelTask()
        {
            _timer?.StopTimer();
            _taskRequested = new();
        }

        public void RequestOrResetTask()
        {
            if (_taskRequested.CheckAndSetFirstCall)
            {
                if (_action is null) { throw new NullReferenceException("Task is null"); }
                var action2 = ResetAfterTask(_action);
                _timer = new TimerWrapper(_frequency);
                _timer.Elapsed += (sender, e) => action2();
                _timer.AutoReset = false;
                _timer.StartTimer();
            }
            else
            {
                ResetTimer();
            }
        }

        public void RequestTask()
        {
            if (_taskRequested.CheckAndSetFirstCall)
            {
                if (_action is null) { throw new NullReferenceException("Task is null"); }
                var action2 = ResetAfterTask(_action);
                _timer = new TimerWrapper(_frequency);
                _timer.Elapsed += (sender, e) => action2();
                _timer.AutoReset = false;
                _timer.StartTimer();
            }
        }

        public void RequestTask(Func<Task> task)
        {
            if (_taskRequested.CheckAndSetFirstCall)
            {
                if (task is null) { throw new NullReferenceException("Task is null"); }
                var task2 = ResetAfterTask(task);
                _timer = new TimerWrapper(_frequency);
                _timer.Elapsed += async (sender, e) => await task2();
                _timer.AutoReset = false;
                _timer.StartTimer();
            }
        }

        private Func<Task> ResetAfterTask(Func<Task> task)
        {
            return async () =>
            {
                await task();
                await Task.Yield();
                _taskRequested = new();
            };
        }
    }
}
