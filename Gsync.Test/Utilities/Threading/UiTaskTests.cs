using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gsync.Utilities.HelperClasses;
using Gsync.Utilities.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Gsync.Test
{
    [TestClass]
    public class UiTaskTests
    {
        private static SynchronizationContext _syncContext;
        private static int _uiThreadId;
        private UiTask _uiTask;
        private static Form _hiddenForm;

        [ClassInitialize]
        public static void ClassInit(TestContext context)
        {
            // Start a WinForms message loop to ensure we are on the UI thread
            var uiThreadReady = new ManualResetEvent(false);

            Thread uiThread = new Thread(() =>
            {
                // Create a hidden form to initialize the WindowsFormsSynchronizationContext
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // This will install the WindowsFormsSynchronizationContext for this thread
                //var form = new Form();
                _hiddenForm = new Form();

                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                //_syncContext = SynchronizationContext.Current;
                //_uiThreadId = Thread.CurrentThread.ManagedThreadId;
                uiThreadReady.Set();

                // Run a minimal message loop
                //Application.Run();               
                Application.Run(_hiddenForm);

            });

            uiThread.SetApartmentState(ApartmentState.STA);
            uiThread.IsBackground = true;
            uiThread.Start();

            // Wait for the UI thread to be ready and context installed
            uiThreadReady.WaitOne();
        }

        [TestInitialize]
        public void Setup()
        {
            Console.SetOut(new DebugTextWriter()); // Redirect Console output to Debug for testing
            //_syncContext = new TestSyncContext(_uiThreadId);            
            _hiddenForm.Invoke(() =>
            {
                _syncContext = SynchronizationContext.Current;
                _uiThreadId = Thread.CurrentThread.ManagedThreadId;
            }); // Ensure we are on the UI thread

            _uiTask = new UiTask(_syncContext, _uiThreadId);
        }

        [TestMethod]
        public async Task Run_Action_ExecutesOnContext()
        {
            int threadId = Thread.CurrentThread.ManagedThreadId;            
            
            if (threadId == _uiThreadId)
            {
                throw new InvalidOperationException("This test must not run on the UI thread initially.");
            }
            else
            {
                await _uiTask.Run(() => threadId = Thread.CurrentThread.ManagedThreadId);
            }

            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_ActionT_ExecutesWithArg()
        {
            int threadId = -1;
            int result = 0;
            await _uiTask.Run<int>(x => { result = x; threadId = Thread.CurrentThread.ManagedThreadId; }, 42);
            Assert.AreEqual(42, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_ActionT1T2_ExecutesWithArgs()
        {
            int threadId = -1;
            int sum = 0;
            await _uiTask.Run<int, int>((a, b) => { sum = a + b; threadId = Thread.CurrentThread.ManagedThreadId; }, 1, 2);
            Assert.AreEqual(3, sum);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_ActionT1T2T3_ExecutesWithArgs()
        {
            int threadId = -1;
            string concat = null;
            await _uiTask.Run<string, string, string>((a, b, c) => { concat = a + b + c; threadId = Thread.CurrentThread.ManagedThreadId; }, "A", "B", "C");
            Assert.AreEqual("ABC", concat);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_ActionT1T2T3T4_ExecutesWithArgs()
        {
            int threadId = -1;
            int result = 0;
            await _uiTask.Run<int, int, int, int>((a, b, c, d) => { result = a + b + c + d; threadId = Thread.CurrentThread.ManagedThreadId; }, 1, 2, 3, 4);
            Assert.AreEqual(10, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_ActionT1T2T3T4T5_ExecutesWithArgs()
        {
            int threadId = -1;
            int result = 0;
            await _uiTask.Run<int, int, int, int, int>((a, b, c, d, e) => { result = a + b + c + d + e; threadId = Thread.CurrentThread.ManagedThreadId; }, 1, 2, 3, 4, 5);
            Assert.AreEqual(15, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_ActionT1T2T3T4T5T6_ExecutesWithArgs()
        {
            int threadId = -1;
            int result = 0;
            await _uiTask.Run<int, int, int, int, int, int>((a, b, c, d, e, f) => { result = a + b + c + d + e + f; threadId = Thread.CurrentThread.ManagedThreadId; }, 1, 2, 3, 4, 5, 6);
            Assert.AreEqual(21, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_ActionT1T2T3T4T5T6T7_ExecutesWithArgs()
        {
            int threadId = -1;
            int result = 0;
            await _uiTask.Run<int, int, int, int, int, int, int>((a, b, c, d, e, f, g) => { result = a + b + c + d + e + f + g; threadId = Thread.CurrentThread.ManagedThreadId; }, 1, 2, 3, 4, 5, 6, 7);
            Assert.AreEqual(28, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_FuncTinTout_ReturnsResult()
        {
            int threadId = -1;
            var task = _uiTask.Run<int, string>(x => { threadId = Thread.CurrentThread.ManagedThreadId; return (x * 2).ToString(); }, 21);
            var result = await task;
            Assert.AreEqual("42", result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_FuncT1T2Tout_ReturnsResult()
        {
            int threadId = -1;
            var task = _uiTask.Run<int, int, int>((a, b) => { threadId = Thread.CurrentThread.ManagedThreadId; return a + b; }, 10, 32);
            var result = await task;
            Assert.AreEqual(42, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_FuncT1T2T3Tout_ReturnsResult()
        {
            int threadId = -1;
            var task = _uiTask.Run<int, int, int, int>((a, b, c) => { threadId = Thread.CurrentThread.ManagedThreadId; return a * b * c; }, 2, 3, 7);
            var result = await task;
            Assert.AreEqual(42, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_FuncT1T2T3T4Tout_ReturnsResult()
        {
            int threadId = -1;
            var task = _uiTask.Run<int, int, int, int, int>((a, b, c, d) => { threadId = Thread.CurrentThread.ManagedThreadId; return a + b + c + d; }, 10, 10, 10, 12);
            var result = await task;
            Assert.AreEqual(42, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_FuncT1T2T3T4T5Tout_ReturnsResult()
        {
            int threadId = -1;
            var task = _uiTask.Run<int, int, int, int, int, int>((a, b, c, d, e) => { threadId = Thread.CurrentThread.ManagedThreadId; return a + b + c + d + e; }, 10, 10, 10, 10, 2);
            var result = await task;
            Assert.AreEqual(42, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        public async Task Run_FuncT1T2T3T4T5T6Tout_ReturnsResult()
        {
            int threadId = -1;
            var task = _uiTask.Run<int, int, int, int, int, int, int>((a, b, c, d, e, f) => { threadId = Thread.CurrentThread.ManagedThreadId; return a + b + c + d + e + f; }, 10, 10, 10, 10, 1, 1);
            var result = await task;
            Assert.AreEqual(42, result);
            Assert.AreEqual(_uiThreadId, threadId);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_NullContext_Throws()
        {
            var _ = new UiTask(null, 1);
        }

        // Helper SynchronizationContext for testing
        
        
    }
}