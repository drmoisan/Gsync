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
        }
        public SynchronizationContext UiContext { get; protected set; }
        public int UiThreadId { get; protected set; }

        public bool IsCurrentThread => Thread.CurrentThread.ManagedThreadId == UiThreadId;
        
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

        public Task RunOnContextAsync(SynchronizationContext context, Action action)
        {
            var tcs = new TaskCompletionSource<object>();

            context.Post(_ =>
            {
                try
                {
                    action();
                    tcs.SetResult(null); // Success
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex); // Propagate error
                }
            }, null);

            return tcs.Task;
        }

    }
}
