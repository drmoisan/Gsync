using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Gsync.Utilities.Threading
{
    /// <summary>
    /// Thread safe enter once into a code block:
    /// the first call to CheckAndSetFirstCall returns always true,
    /// all subsequent call return false.
    /// https://www.codeproject.com/Tips/375559/Implement-Thread-Safe-One-shot-Bool-Flag-with-Inte
    /// </summary>
    public class ThreadSafeSingleShotGuard
    {
        private static int NOTCALLED = 0;
        private static int CALLED = 1;
        private int _state = NOTCALLED;
        /// <summary>Explicit call to check and set if this is the first call</summary>
        public bool CheckAndSetFirstCall
        { get { return Interlocked.Exchange(ref _state, CALLED) == NOTCALLED; } }
    }

    
}
