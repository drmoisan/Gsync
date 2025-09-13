using log4net.Repository.Hierarchy;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

using Gsync.Utilities.Interfaces;

namespace Gsync.Utilities.HelperClasses
{
    /// <summary>
    /// Producer/Consumer pattern for writing items of type <typeparamref name="T"/> to disk 
    /// with a delegate on a regular interval
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class TimedDiskWriter<T>
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// Constructor for <see cref="TimedDiskWriter{T}"/>
        /// </summary>
        public TimedDiskWriter() 
        {
            Config = new Configuration(20, TimeSpan.FromMilliseconds(500));
            Config.PropertyChanged += Configuration_PropertyChanged;
        }

        /// <param name="writeInterval">TimeSpan representing frequency of writing to disk</param>
        /// <param name="diskWriter">Delegate to write Queue to disk</param>
        /// <inheritdoc cref="TimedDiskWriter()"/>
        public TimedDiskWriter(TimeSpan writeInterval, Action<IEnumerable<T>> diskWriter)
        {
            Config = new Configuration(20, writeInterval);
            Config.PropertyChanged += Configuration_PropertyChanged;
            DiskWriter = diskWriter;
        }

        /// <param name="milliseconds">Integer representing interval in milliseconds of writing to disk</param>
        /// <param name="diskWriter">Delegate to write Queue to disk</param>
        /// <inheritdoc cref="TimedDiskWriter()"/>
        public TimedDiskWriter(int milliseconds, Action<IEnumerable<T>> diskWriter)
        {
            Config = new Configuration(20, TimeSpan.FromMilliseconds(milliseconds));
            Config.PropertyChanged += Configuration_PropertyChanged;
            DiskWriter = diskWriter;
        }

        #region Public Properties
                
        /// <inheritdoc cref="Configuration"/>
        public virtual Configuration Config { get => _config; private set => _config = value; }
        private Configuration _config;

        /// <summary>
        /// Delegate to write an <see cref="IEnumerable{T}">IEnumerable&lt;T&gt;</see> to disk
        /// </summary>
        public Action<IEnumerable<T>> DiskWriter { get => _diskWriter; set => _diskWriter = value; }
        private Action<IEnumerable<T>> _diskWriter;

        /// <summary>
        /// Queue of items to be written to disk
        /// </summary>
        public BlockingCollection<T> Queue { get => _queue; internal set => _queue = value; }
        private BlockingCollection<T> _queue = new(new ConcurrentQueue<T>());

        private ITimerWrapper _timer;
        internal ITimerWrapper Timer { get => _timer; set => _timer = value; }

        #endregion

        #region Public Producer / Consumer Methods

        public void Enqueue(T item)
        {
            if (DiskWriter is null) { logger.Warn($"{nameof(TimedDiskWriter<T>)} is Enqueuing items with no function to write the items to disk"); }
            if (!TimerActive)
            {
                if (!TryStartTimer()) { logger.Warn($"{nameof(TimedDiskWriter<T>)} is Enqueuing items and is unable to start the timer to write the items to disk"); }
            }
            CancellationTokenSource cts = new();
            var token = cts.Token;
            var success = false;

            do
            {
                try
                {
                    success = Queue.TryAdd(item, Config.TryAddTimeout, token);
                }
                catch (OperationCanceledException)
                {
                    if (token.IsCancellationRequested) { break; }
                    else
                    {
                        //logger.Debug($"Timeout adding {item}");
                    }
                }
            } while (!success);
        }

        public async Task EnqueueAsync(T item, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            if (DiskWriter is null) { logger.Warn($"{nameof(TimedDiskWriter<T>)} is Enqueuing items with no function to write the items to disk"); }
            if (!TimerActive)
            {
                if (!TryStartTimer()) { logger.Warn($"{nameof(TimedDiskWriter<T>)} is Enqueuing items and is unable to start the timer to write the items to disk"); }
            }

            var success = false;

            do
            {
                try
                {
                    success = Queue.TryAdd(item, Config.TryAddTimeout, token);
                }
                catch (OperationCanceledException)
                {
                    if (token.IsCancellationRequested) { break; }
                    else
                    {
                        //logger.Debug($"Timeout adding {item}");
                        await Task.Delay(Config.TryAddTimeout);
                    }
                }
            } while (!success);
        }

        public virtual bool TimerActive => _timer is not null && _timer.Enabled;

        private int _emptyQueueChecks = 0;
        /// <summary>
        /// Callback function for the <seealso cref="System.Timers.Timer">Timer</seealso> and 
        /// "Consumer" for <see cref="Queue"/> which invokes the <see cref="DiskWriter"/>
        /// </summary>
        /// <param name="sender">Timer object</param>
        /// <param name="e">Elapsed event arguments</param>
        internal void OnTimedEvent(object sender, TimeElapsedEventArgs e)
        {
            //var items = Queue.GetConsumingEnumerable();
            var items = new List<T>();
            while (Queue.TryTake(out var item))
            {
                items.Add(item);
            }
            
            if (items.Any())
            {
                DiskWriter(items);
            }
            else 
            { 
                Interlocked.Increment(ref _emptyQueueChecks);
                if (_emptyQueueChecks > 4)
                {
                    StopTimer();
                    _emptyQueueChecks = 0;
                }
            }
        }
        
        /// <summary>
        /// Starts the timer to invoke the <see cref="DiskWriter"/>
        /// </summary>
        public virtual void StartTimer()
        {
            if (DiskWriter is null) 
            { 
                throw new InvalidOperationException($"{nameof(TimedDiskWriter<T>)} is " +
                    $"attempting to start the timer with no action assigned to " +
                    $"the callback {nameof(DiskWriter)} ");
            }
            else
            {
                //_timer = new System.Timers.Timer(Config.WriteInterval.TotalMilliseconds);
                _timer = new TimerWrapper(Config.WriteInterval);
                _timer.Elapsed += OnTimedEvent;
                _timer.AutoReset = true;
                _timer.Enabled = true;
            }
        }

        /// <summary>
        /// Attempts to start the timer to invoke the <see cref="DiskWriter"/>
        /// </summary>
        /// <returns>true if successful. false if unable to start</returns>
        internal virtual bool TryStartTimer()
        {
            try
            {
                StartTimer();
                return true;
            }
            catch (InvalidOperationException e)
            {
                logger.Warn($"Aborting operation {nameof(TryStartTimer)}. {e.Message}");
                return false;
            }
        }

        /// <summary>
        /// Stops the timer from invoking the <see cref="DiskWriter"/> and disposes of the timer
        /// </summary>
        public virtual void StopTimer()
        {
            _timer?.StopTimer();
            _timer?.Dispose();
        }

        #endregion

        #region Configuration Class and Event Handler

        /// <summary>
        /// Event Handler for the <see cref="Configuration.PropertyChanged"/> event.
        /// Restarts timer with new configuration
        /// </summary>
        /// <param name="sender">Reference to PropertyChanged event</param>
        /// <param name="e">Details about which property triggered the event</param>
        public void Configuration_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(Config.WriteInterval))
            {
                StopTimer();
                TryStartTimer();
            }
        }

        /// <summary>
        /// Holds configuration settings for the <see cref="TimedDiskWriter{T}"/> 
        /// class and notifies when the properties change
        /// </summary>
        public class Configuration : INotifyPropertyChanged
        {
            /// <summary>
            /// Constructor for the <see cref="Configuration"/> class
            /// </summary>
            public Configuration() { }

            /// <summary>
            /// Constructor for the <see cref="Configuration"/> class that accepts two parameters:
            /// <list type="bullet">
            /// <item>
            /// <term>tryAddTimeout</term>
            /// <description>Timeout interval in milliseconds for the 
            /// <seealso cref="BlockingCollection{T}.TryAdd(T)"/> method</description>
            /// </item>
            /// /// <item>
            /// <term>writeInterval</term>
            /// <description>Frequency with which to write to disk</description>
            /// </item>
            /// </list>
            /// </summary>
            /// <param name="tryAddTimeout">Timeout interval in milliseconds for the 
            /// <seealso cref="BlockingCollection{T}.TryAdd(T)"/> method 
            /// within <seealso cref="EnqueueAsync(T, CancellationToken)"/></param>
            /// <param name="writeInterval"></param>
            public Configuration(int tryAddTimeout, TimeSpan writeInterval)
            {
                _tryAddTimeout = tryAddTimeout;
                _writeInterval = writeInterval;
            }

            /// <summary>
            /// Timeout in milliseconds for adding to the queue
            /// </summary>
            public int TryAddTimeout { get => _tryAddTimeout; set => _tryAddTimeout = value; }
            private int _tryAddTimeout = 20;

            /// <summary>
            /// Frequency with which to write to disk
            /// </summary>
            public TimeSpan WriteInterval 
            { 
                get => _writeInterval;
                set 
                { 
                    _writeInterval = value;
                    NotifyPropertyChanged();
                } 
            }
            private TimeSpan _writeInterval = TimeSpan.FromMinutes(5);

            /// <summary>
            /// Helper method to raise the <see cref="PropertyChanged"/> event
            /// </summary>
            /// <param name="propertyName">Argument that specifies the property that changed. 
            /// If left blank, it is inferred from the caller member name</param>
            public void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }

            /// <summary>
            /// When the <see cref="WriteInterval"/> is changed, the <see cref="PropertyChanged"/> event is raised
            /// </summary>
            public event PropertyChangedEventHandler PropertyChanged;
        }

        #endregion
    }


}
