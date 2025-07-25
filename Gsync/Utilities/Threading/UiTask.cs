using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Gsync.Utilities.Threading
{
    public class UiTask
    {
        public UiTask(SynchronizationContext uiContext, int uiThreadId)
        {
            UiContext = uiContext ?? throw new ArgumentNullException(nameof(uiContext), "SynchronizationContext cannot be null.");
            UiThreadId = uiThreadId;
        }

        public SynchronizationContext UiContext { get; protected set; }
        public int UiThreadId { get; protected set; }

        public Task Run(Action action)
        {
            var tcs = new TaskCompletionSource<object>();

            UiContext.Post(_ =>
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

        public Task Run<T>(Action<T> action, T arg)
        {
            var tcs = new TaskCompletionSource<object>();

            UiContext.Post(_ =>
            {
                try
                {
                    action(arg);
                    tcs.SetResult(null); // Success
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex); // Propagate error
                }
            }, null);

            return tcs.Task;
        }

        public Task Run<T1, T2>(Action<T1, T2> action, T1 arg1, T2 arg2)
        {
            var tcs = new TaskCompletionSource<object>();

            UiContext.Post(_ =>
            {
                try
                {
                    action(arg1, arg2);
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task Run<T1, T2, T3>(Action<T1, T2, T3> action, T1 arg1, T2 arg2, T3 arg3)
        {
            var tcs = new TaskCompletionSource<object>();

            UiContext.Post(_ =>
            {
                try
                {
                    action(arg1, arg2, arg3);
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task Run<T1, T2, T3, T4>(Action<T1, T2, T3, T4> action, T1 arg1, T2 arg2, T3 arg3, T4 arg4)
        {
            var tcs = new TaskCompletionSource<object>();

            UiContext.Post(_ =>
            {
                try
                {
                    action(arg1, arg2, arg3, arg4);
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task Run<T1, T2, T3, T4, T5>(Action<T1, T2, T3, T4, T5> action, T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5 arg5)
        {
            var tcs = new TaskCompletionSource<object>();

            UiContext.Post(_ =>
            {
                try
                {
                    action(arg1, arg2, arg3, arg4, arg5);
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task Run<T1, T2, T3, T4, T5, T6>(Action<T1, T2, T3, T4, T5, T6> action, T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5 arg5, T6 arg6)
        {
            var tcs = new TaskCompletionSource<object>();

            UiContext.Post(_ =>
            {
                try
                {
                    action(arg1, arg2, arg3, arg4, arg5, arg6);
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task Run<T1, T2, T3, T4, T5, T6, T7>(Action<T1, T2, T3, T4, T5, T6, T7> action, T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5 arg5, T6 arg6, T7 arg7)
        {
            var tcs = new TaskCompletionSource<object>();

            UiContext.Post(_ =>
            {
                try
                {
                    action(arg1, arg2, arg3, arg4, arg5, arg6, arg7);
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task<TOut> Run<TIn, TOut>(Func<TIn, TOut> func, TIn arg)
        {
            var tcs = new TaskCompletionSource<TOut>();

            UiContext.Post(_ =>
            {
                try
                {
                    var result = func(arg);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task<TOut> Run<T1, T2, TOut>(Func<T1, T2, TOut> func, T1 arg1, T2 arg2)
        {
            var tcs = new TaskCompletionSource<TOut>();

            UiContext.Post(_ =>
            {
                try
                {
                    var result = func(arg1, arg2);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task<TOut> Run<T1, T2, T3, TOut>(Func<T1, T2, T3, TOut> func, T1 arg1, T2 arg2, T3 arg3)
        {
            var tcs = new TaskCompletionSource<TOut>();

            UiContext.Post(_ =>
            {
                try
                {
                    var result = func(arg1, arg2, arg3);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task<TOut> Run<T1, T2, T3, T4, TOut>(Func<T1, T2, T3, T4, TOut> func, T1 arg1, T2 arg2, T3 arg3, T4 arg4)
        {
            var tcs = new TaskCompletionSource<TOut>();

            UiContext.Post(_ =>
            {
                try
                {
                    var result = func(arg1, arg2, arg3, arg4);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task<TOut> Run<T1, T2, T3, T4, T5, TOut>(Func<T1, T2, T3, T4, T5, TOut> func, T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5 arg5)
        {
            var tcs = new TaskCompletionSource<TOut>();

            UiContext.Post(_ =>
            {
                try
                {
                    var result = func(arg1, arg2, arg3, arg4, arg5);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }

        public Task<TOut> Run<T1, T2, T3, T4, T5, T6, TOut>(Func<T1, T2, T3, T4, T5, T6, TOut> func, T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5 arg5, T6 arg6)
        {
            var tcs = new TaskCompletionSource<TOut>();

            UiContext.Post(_ =>
            {
                try
                {
                    var result = func(arg1, arg2, arg3, arg4, arg5, arg6);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return tcs.Task;
        }
    }
}
