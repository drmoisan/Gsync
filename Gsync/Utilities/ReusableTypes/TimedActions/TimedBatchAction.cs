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
    public class TimedBatchAction
    {
        public TimedBatchAction(TimeSpan frequency)
        {
            _frequency = frequency;
        }

        public TimedBatchAction(TimeSpan frequency, System.Action action)
        {
            _frequency = frequency;
            _action = action;
        }

        public void SetAction(System.Action action)
        {
            Interlocked.CompareExchange(ref _action, action, null);
        }
        private System.Action _action;
                
        private TimeSpan _frequency;
        private ThreadSafeSingleShotGuard _actionRequested = new();
        private TimerWrapper _timer;
        
        public void ResetTimer()
        {
            _timer?.ResetTimer();
        }

        public void CancelAction() 
        {
            _timer?.StopTimer();
            _actionRequested = new();
        }

        public void RequestAction()
        {
            if (_actionRequested.CheckAndSetFirstCall)
            {
                if (_action is null) { throw new NullReferenceException("Action is null"); }
                var action2 = ResetAfterAction(_action);
                _timer = new TimerWrapper(_frequency);
                _timer.Elapsed += (sender, e) => action2();
                _timer.AutoReset = false;
                _timer.StartTimer();
            }
        }

        public void RequestAction(System.Action action)
        {
            if (_actionRequested.CheckAndSetFirstCall)
            {
                if (action is null) { throw new NullReferenceException("Action is null"); }
                var action2 = ResetAfterAction(action);
                _timer = new TimerWrapper(_frequency);
                _timer.Elapsed += (sender, e) => action2();
                _timer.AutoReset = false;
                _timer.StartTimer();
            }
        }

        private System.Action ResetAfterAction(System.Action action)
        {
            return () =>
            {
                action();
                _actionRequested = new();
            };
        }
    }
}
