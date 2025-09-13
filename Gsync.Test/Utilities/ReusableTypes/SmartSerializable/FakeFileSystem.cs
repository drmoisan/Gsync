using Gsync.Utilities.Interfaces;
using System;
using System.IO;
using System.Windows.Forms;

namespace Gsync.Test.Utilities.ReusableTypes.SmartSerializable
{
    public class FakeFileSystem : IFileSystem
    {
        public bool FileExistsResult = true;
        public string FileContent = "";
        public string WrittenPath;
        public string WrittenContent;
        public bool Exists(string path) => FileExistsResult;
        public string ReadAllText(string path) => FileContent;
        public void WriteAllText(string path, string contents)
        {
            WrittenPath = path;
            WrittenContent = contents;
        }
        public StreamWriter CreateText(string path)
        {
            WrittenPath = path;
            var ms = new MemoryStream();
            var sw = new StreamWriter(ms);
            sw.AutoFlush = true;
            return sw;
        }

        
    }

    public class FakeUserDialog : IUserDialog
    {
        public DialogResult ResultToReturn = DialogResult.Yes;
        public string LastMessage;
        public DialogResult ShowDialog(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            LastMessage = text;
            return ResultToReturn;
        }
    }

    public class FakeTimerFactory : ITimerFactory
    {
        public ITimerWrapper LastTimer;
        public bool ImmediateTimer { get; set; } = false;
        public ITimerWrapper CreateTimer(System.TimeSpan interval)
        {
            var timer = new FakeTimer();
            if (ImmediateTimer) { timer.Immediate = true; }
            LastTimer = timer;
            return LastTimer;
        }
    }

    public class FakeTimer : ITimerWrapper
    {
        public bool Started { get; private set; }
        public bool Disposed { get; private set; }
        public bool Immediate { get; set; } = false;

        public event EventHandler<TimeElapsedEventArgs> Elapsed;

        public double IntervalInMilliseconds { get; set; }
        public TimeSpan Interval { get; set; }
        public bool Enabled { get; set; }
        public bool AutoReset { get; set; }

        public void StartTimer()
        {
            Started = true;
            if (Immediate)
            {
                // Immediately fire the Elapsed event as if the timer expired
                Elapsed?.Invoke(this, new TimeElapsedEventArgs(DateTime.Now));
            }
        }

        public void StopTimer()
        {
            Enabled = false;
        }

        public void ResetTimer() { }

        public void Dispose()
        {
            Disposed = true;
        }
    }


}
