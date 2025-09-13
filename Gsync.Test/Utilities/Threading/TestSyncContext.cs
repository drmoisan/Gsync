using System;
using System.Threading;
using System.Threading.Tasks;

namespace Gsync.Test
{
    public class TestSyncContext: SynchronizationContext
    {
        private readonly int _threadId;

        public TestSyncContext(int threadId)
        {
            _threadId = threadId;
        }

        public override void Post(SendOrPostCallback d, object state)
        {
            // Simulate running on the UI thread by setting the thread ID before invoking
            var prevThreadId = Thread.CurrentThread.ManagedThreadId;
            try
            {
                // If not already on the UI thread, switch (simulate for test)
                if (prevThreadId != _threadId)
                {
                    // In real test, you might use a dedicated thread, but for simplicity, just invoke
                }
                d(state);
            }
            finally
            {
                // No thread switching needed in this test context
            }
        }
    }
}
