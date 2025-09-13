using System.Threading;
using Gsync.Utilities.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Gsync.Test.Utilities.Threading
{
    [TestClass]
    public class ThreadSafeSingleShotGuardTests
    {
        [TestMethod]
        public void CheckAndSetFirstCall_ReturnsTrue_OnFirstCall()
        {
            var guard = new ThreadSafeSingleShotGuard();

            var result = guard.CheckAndSetFirstCall;

            Assert.IsTrue(result, "First call should return true.");
        }

        [TestMethod]
        public void CheckAndSetFirstCall_ReturnsFalse_OnSubsequentCalls()
        {
            var guard = new ThreadSafeSingleShotGuard();

            var first = guard.CheckAndSetFirstCall;
            var second = guard.CheckAndSetFirstCall;
            var third = guard.CheckAndSetFirstCall;

            Assert.IsTrue(first, "First call should return true.");
            Assert.IsFalse(second, "Second call should return false.");
            Assert.IsFalse(third, "Third call should return false.");
        }

        [TestMethod]
        public void CheckAndSetFirstCall_IsThreadSafe()
        {
            var guard = new ThreadSafeSingleShotGuard();
            var results = new bool[10];
            var threads = new Thread[10];

            for (int i = 0; i < 10; i++)
            {
                int idx = i;
                threads[i] = new Thread(() => { results[idx] = guard.CheckAndSetFirstCall; });
            }

            foreach (var t in threads) t.Start();
            foreach (var t in threads) t.Join();

            int trueCount = 0;
            int falseCount = 0;
            foreach (var r in results)
            {
                if (r) trueCount++;
                else falseCount++;
            }

            Assert.AreEqual(1, trueCount, "Only one thread should get true.");
            Assert.AreEqual(9, falseCount, "All other threads should get false.");
        }
    }
}