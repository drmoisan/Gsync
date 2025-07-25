using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Gsync.Utilities.Threading
{
    public class UiWrapper
    {
        internal UiWrapper() { }

        public UiWrapper(SynchronizationContext context, int threadId)
        {
            UiContext = context ?? throw new ArgumentNullException(nameof(context), "SynchronizationContext cannot be null.");
            UiThreadId = threadId;
            this.Task = new UiTask(context, threadId);
        }

        public SynchronizationContext UiContext { get; protected set; }
        public int UiThreadId { get; protected set; }

        public bool IsCurrentThread => Thread.CurrentThread.ManagedThreadId == UiThreadId;
        
        public UiTask Task { get; protected set; }
        
        public void Invoke(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action), "Action cannot be null.");
            if (UiContext == null) throw new InvalidOperationException("SynchronizationContext is not set.");
            UiContext.Send(_ => action(), null);            
        }

        public void InvokeSafe(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action), "Action cannot be null.");
            if (UiContext == null) throw new InvalidOperationException("SynchronizationContext is not set.");
            if (IsCurrentThread)
            {
                action();
            }
            else
            {
                UiContext.Send(_ => action(), null);
            }
        }

    }
}
