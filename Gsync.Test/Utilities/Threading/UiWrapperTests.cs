using System;
using System.Threading;
using Gsync.Utilities.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Gsync.Test.Utilities.Threading
{
    [TestClass]
    public class UiWrapperTests
    {
        private SynchronizationContext _syncContext;
        private int _uiThreadId;
        private UiWrapper _uiWrapper;

        private class TestSyncContext : SynchronizationContext
        {
            public int SendCallCount { get; private set; }
            public override void Send(SendOrPostCallback d, object state)
            {
                SendCallCount++;
                d(state);
            }
        }

        private class TestableUiWrapper : UiWrapper
        {
            public TestableUiWrapper(SynchronizationContext context, int threadId)
                : base(context, threadId) { }

            public void SetUiContext(SynchronizationContext context)
            {
                this.UiContext = context;
            }
        }

        [TestInitialize]
        public void Setup()
        {
            _uiThreadId = Thread.CurrentThread.ManagedThreadId;
            _syncContext = new TestSyncContext();
            _uiWrapper = new UiWrapper(_syncContext, _uiThreadId);
        }

        [TestMethod]
        public void Constructor_Parameterless_CreatesInstance()
        {
            var wrapper = new UiWrapper();
            Assert.IsNotNull(wrapper);
            Assert.IsNull(wrapper.UiContext);
            Assert.AreEqual(0, wrapper.UiThreadId);
            Assert.IsNull(wrapper.Task);
        }

        [TestMethod]
        public void Constructor_ValidArgs_InitializesProperties()
        {
            Assert.AreEqual(_syncContext, _uiWrapper.UiContext);
            Assert.AreEqual(_uiThreadId, _uiWrapper.UiThreadId);
            Assert.IsNotNull(_uiWrapper.Task);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullContext_Throws()
        {
            var _ = new UiWrapper(null, 1);
        }

        [TestMethod]
        public void IsCurrentThread_ReturnsTrueForUiThread()
        {
            Assert.IsTrue(_uiWrapper.IsCurrentThread);
        }

        [TestMethod]
        public void IsCurrentThread_ReturnsFalseForOtherThread()
        {
            int result = -1;
            var thread = new Thread(() =>
            {
                var wrapper = new UiWrapper(_syncContext, _uiThreadId);
                result = wrapper.IsCurrentThread ? 1 : 0;
            });
            thread.Start();
            thread.Join();
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void Invoke_ExecutesActionOnContext()
        {
            var testContext = (TestSyncContext)_syncContext;
            bool called = false;
            _uiWrapper.Invoke(() => called = true);
            Assert.IsTrue(called);
            Assert.AreEqual(1, testContext.SendCallCount);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Invoke_NullAction_Throws()
        {
            _uiWrapper.Invoke(null);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void Invoke_NullContext_Throws()
        {
            var wrapper = new TestableUiWrapper(_syncContext, _uiThreadId);
            wrapper.SetUiContext(null);
            wrapper.Invoke(() => { });
        }

        [TestMethod]
        public void InvokeSafe_OnUiThread_ExecutesDirectly()
        {
            bool called = false;
            _uiWrapper.InvokeSafe(() => called = true);
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void InvokeSafe_OnOtherThread_UsesContext()
        {
            var testContext = new TestSyncContext();
            var wrapper = new UiWrapper(testContext, _uiThreadId);
            bool called = false;
            var thread = new Thread(() =>
            {
                wrapper.InvokeSafe(() => called = true);
            });
            thread.Start();
            thread.Join();
            Assert.IsTrue(called);
            Assert.AreEqual(1, testContext.SendCallCount);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void InvokeSafe_NullAction_Throws()
        {
            _uiWrapper.InvokeSafe(null);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void InvokeSafe_NullContext_Throws()
        {
            var wrapper = new TestableUiWrapper(_syncContext, _uiThreadId);
            wrapper.SetUiContext(null);
            wrapper.InvokeSafe(() => { });
        }
    }
}