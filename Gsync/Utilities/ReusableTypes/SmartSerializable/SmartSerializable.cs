using Gsync.Utilities.Extensions;
using Gsync.Utilities.HelperClasses;
using Gsync.Utilities.Interfaces;
using Gsync.Utilities.Threading;
using Newtonsoft.Json;
using System;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gsync.Utilities.ReusableTypes
{
    public class SmartSerializable<T> : ISmartSerializable<T>, ISmartSerializableLoader<T>
        where T : class, ISmartSerializable<T>, new()
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        // Dependencies to be injected (via test, or assigned for production)
        public IFileSystem FileSystem { get; set; }
        public IUserDialog UserDialog { get; set; }
        public ITimerFactory TimerFactory { get; set; }

        public SmartSerializable()
        {
            _parent = null;
            Config = new NewSmartSerializableConfig();
        }

        public SmartSerializable(T parent)
        {
            _parent = parent;
            Config = new NewSmartSerializableConfig();
        }

        protected T _parent;

        private ISmartSerializableConfig _config = new NewSmartSerializableConfig();
        public ISmartSerializableConfig Config
        {
            get => _config;
            set
            {
                if (_config is not null)
                    _config.PropertyChanged -= Config_PropertyChanged;
                _config = value;
                if (_config is not null)
                    _config.PropertyChanged += Config_PropertyChanged;
            }
        }

        public string Name { get; set; }

        private void Config_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            Notify(e.PropertyName);
        }

        public void Notify([System.Runtime.CompilerServices.CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        #region Deserialization

        protected T CreateEmpty(DialogResult response, FilePathHelper disk, JsonSerializerSettings settings, Func<T> altLoader = null)
        {
            if (response == DialogResult.Yes)
            {
                var instance = altLoader is null ? new T() : altLoader();
                if (instance == null)
                    throw new InvalidOperationException($"{nameof(altLoader)} returned null instance.");
                instance.Config.JsonSettings = settings;
                instance.Serialize(disk.FilePath);
                return instance;
            }
            else
            {
                throw new InvalidOperationException(
                    $"Must have an instance of {typeof(T)} or create one to continue executing");
            }
        }


        protected DialogResult AskUser(bool askUserOnError, string messageText)
        {
            if (askUserOnError && UserDialog is not null)
            {                
                return UserDialog.ShowDialog(messageText, "Error",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);                
            }
            else
            {
                return DialogResult.Yes;
            }
        }

        /// <summary>
        /// Deserializes an instance from the specified file.
        /// This overload does NOT prompt the user on error—if the file is missing or corrupted,
        /// it silently creates a new instance instead of asking the user or throwing an exception.
        /// To enable user prompting and error handling, use an overload that exposes the 'askUserOnError' parameter.
        /// </summary>
        /// <param name="fileName">The name of the file to deserialize.</param>
        /// <param name="folderPath">The folder path where the file is located.</param>
        /// <returns>An instance of T, either loaded from disk or newly created if not found/corrupt.</returns>
        public T Deserialize(string fileName, string folderPath)
            => Deserialize(fileName, folderPath, false);

        public T Deserialize(string fileName, string folderPath, bool askUserOnError)
        {
            var disk = new FilePathHelper(fileName, folderPath);
            var settings = GetDefaultSettings();
            return Deserialize(disk, askUserOnError, settings);
        }

        public T Deserialize(string fileName, string folderPath, bool askUserOnError, JsonSerializerSettings settings)
        {
            var disk = new FilePathHelper(fileName, folderPath);
            return Deserialize(disk, askUserOnError, settings);
        }

        public T Deserialize<U>(ISmartSerializableLoader<U> loader) where U : class, ISmartSerializable<U>, new()
        {
            try
            {
                var disk = loader.ThrowIfNull().Config.ThrowIfNull().Disk.ThrowIfNull();
                var settings = loader.Config.JsonSettings.ThrowIfNull();
                T instance = DeserializeJson(loader.Config.Disk, loader.Config.JsonSettings);
                if (instance is not null) { instance.Config.CopyFrom(loader.Config, true); }
                return instance;
            }
            catch (ArgumentNullException e)
            {
                logger.Error(e.Message);
                throw;
            }
        }

        public T Deserialize<U>(ISmartSerializableLoader<U> loader, bool askUserOnError, Func<T> altLoader)
            where U : class, ISmartSerializable<U>, new()
        {
            var disk = loader.ThrowIfNull().Config.ThrowIfNull().Disk.ThrowIfNull();
            var settings = loader.Config.JsonSettings.ThrowIfNull();
            bool writeInstance = false;
            T instance = default;

            try
            {
                instance = DeserializeJson(loader.Config.Disk, loader.Config.JsonSettings);
                if (instance is null)
                {
                    throw new InvalidOperationException($"{disk.FilePath} deserialized to null.");
                }
            }
            catch (FileNotFoundException e)
            {
                logger.Error(e.Message);
                var response = AskUser(askUserOnError,
                    $"{disk.FilePath} not found. Need an instance of {typeof(T)} to " +
                    $"continue. Create a new dictionary or abort execution?");
                if (response == DialogResult.Yes)
                {
                    instance = CreateEmpty(response, disk, settings, altLoader);
                    writeInstance = true;
                }
                else
                {
                    throw new InvalidOperationException(
                        $"Cannot continue execution without a valid instance of {typeof(T)}.");
                }                
            }
            catch (System.Exception e)
            {
                logger.Error($"Error! {e.Message}");
                var response = AskUser(askUserOnError,
                    $"{disk.FilePath} encountered a problem. \n{e.Message}\n" +
                    $"Need a dictionary to continue. Create a new dictionary or abort execution?");
                instance = CreateEmpty(response, disk, settings, altLoader);
                writeInstance = true;
            }
            instance.Config.CopyFrom(loader.Config, true);

            if (writeInstance)
            {
                instance.Serialize();
            }

            return instance;
        }

        protected T Deserialize(FilePathHelper disk, bool askUserOnError, JsonSerializerSettings settings)
        {
            bool writeInstance = false;
            T instance;
            DialogResult response;

            try
            {
                instance = DeserializeJson(disk, settings);
                if (instance is null)
                {
                    throw new InvalidOperationException($"{disk.FilePath} deserialized to null.");
                }
            }
            catch (FileNotFoundException e)
            {
                logger.Error(e.Message);
                response = AskUser(askUserOnError,
                    $"{disk.FilePath} not found. Need an instance of {typeof(T)} to " +
                    $"continue. Create a new dictionary or abort execution?");
                instance = CreateEmpty(response, disk, settings);
                writeInstance = true;
            }
            catch (System.Exception e)
            {
                logger.Error($"Error! {e.Message}");
                response = AskUser(askUserOnError,
                    $"{disk.FilePath} encountered a problem. \n{e.Message}\n" +
                    $"Need a dictionary to continue. Create a new dictionary or abort execution?");
                if (response == DialogResult.Yes) 
                { 
                    instance = CreateEmpty(response, disk, settings);
                    writeInstance = true;
                }
                else 
                {                     
                    throw new InvalidOperationException(
                        $"Cannot continue execution without a valid instance of {typeof(T)}.");
                }
            }

            instance.Config.Disk.FilePath = disk.FilePath;

            if (writeInstance)
            {
                instance.Serialize();
            }
            return instance;
        }

        public async Task<T> DeserializeAsync<U>(ISmartSerializableLoader<U> loader) where U : class, ISmartSerializable<U>, new()
        {
            return await Task.Run(() => Deserialize(loader));
        }

        public async Task<T> DeserializeAsync<U>(ISmartSerializableLoader<U> loader, bool askUserOnError) where U : class, ISmartSerializable<U>, new()
        {
            return await Task.Run(() => Deserialize(loader, askUserOnError, null));
        }

        public async Task<T> DeserializeAsync<U>(ISmartSerializableLoader<U> loader, bool askUserOnError, Func<T> altLoader) where U : class, ISmartSerializable<U>, new()
        {
            return await Task.Run(() => Deserialize(loader, askUserOnError, altLoader));
        }

        protected T DeserializeJson(FilePathHelper disk, JsonSerializerSettings settings)
        {
            T instance = null;
            //if (!FileSystem.Exists(disk.FilePath)) { return instance; }
            if (!FileSystem.Exists(disk.FilePath)) { throw new FileNotFoundException(disk.FilePath); }
            try
            {
                instance = JsonConvert.DeserializeObject<T>(
                    FileSystem.ReadAllText(disk.FilePath), settings);
            }
            catch (Exception e)
            {
                logger.Error(e.Message, e);
            }
            if (instance is not null) { instance.Config.JsonSettings = settings; }
            return instance;
        }

        public T DeserializeObject(string json, JsonSerializerSettings settings)
        {
            T instance = null;
            try
            {
                instance = JsonConvert.DeserializeObject<T>(json, settings);
            }
            catch (Exception e)
            {
                logger.Error(e.Message, e);
            }
            if (instance is not null)
            {
                instance.Config.JsonSettings = settings.DeepCopy();
            }
            return instance;
        }

        protected T DeserializeJson(FilePathHelper disk)
        {
            var settings = GetDefaultSettings();
            return DeserializeJson(disk, settings);
        }

        #endregion Deserialization

        #region Serialization

        public void Serialize()
        {
            if (!string.IsNullOrEmpty(Config.Disk.FilePath))
            {
                RequestSerialization(Config.Disk.FilePath);
            }
        }

        public void Serialize(string filePath)
        {
            Config.Disk.FilePath = filePath;
            RequestSerialization(filePath);
        }

        protected ReaderWriterLockSlim _readWriteLock = new();

        public static JsonSerializerSettings GetDefaultSettings()
        {
            return new JsonSerializerSettings()
            {
                TypeNameHandling = TypeNameHandling.Auto,
                Formatting = Formatting.Indented
            };
        }

        private Func<string, StreamWriter> _createStreamWriter;
        protected internal Func<string, StreamWriter> CreateStreamWriter
        {
            get => _createStreamWriter ?? FileSystem.CreateText;
            set => _createStreamWriter = value;
        }

        public void SerializeThreadSafe(string filePath)
        {
            _parent.ThrowIfNull($"{nameof(SmartSerializable<T>)}.{nameof(_parent)} is null. It must be linked to the instance it is serializing.");
            if (_readWriteLock.TryEnterWriteLock(-1))
            {
                try
                {
                    using (StreamWriter sw = CreateStreamWriter(filePath))
                    {
                        SerializeToStream(sw);
                        sw.Close();
                    }
                }
                catch (System.Exception e)
                {
                    logger.Error($"Error serializing to {filePath}", e);
                }
                finally
                {
                    _readWriteLock.ExitWriteLock();
                    _serializationRequested = new ThreadSafeSingleShotGuard();
                }
            }
        }

        public string SerializeToString()
        {
            using var memoryStream = new MemoryStream();
            using var streamWriter = new StreamWriter(memoryStream);
            try
            {
                SerializeToStream(streamWriter);
                streamWriter.Flush();
                memoryStream.Position = 0;
            }
            catch (Exception e)
            {
                logger.Error($"Error serializing to string", e);
                return "";
            }
            using var streamReader = new StreamReader(memoryStream);
            return streamReader.ReadToEnd();
        }

        public void SerializeToStream(StreamWriter sw)
        {
            sw.ThrowIfNull();
            var serializer = JsonSerializer.Create(Config.JsonSettings);

            if (Config.JsonSettings.TypeNameHandling == TypeNameHandling.Auto)
            {
                serializer.Serialize(sw, _parent, _parent.GetType());
            }
            else
            {
                serializer.Serialize(sw, _parent);
            }
        }

        private ThreadSafeSingleShotGuard _serializationRequested = new();
        private ITimerWrapper _timer;

        protected void RequestSerialization(string filePath)
        {
            if (_serializationRequested.CheckAndSetFirstCall)
            {
                _timer = TimerFactory.CreateTimer(TimeSpan.FromSeconds(3));
                _timer.Elapsed += (sender, e) => SerializeThreadSafe(filePath);
                _timer.AutoReset = false;
                _timer.StartTimer();
            }
        }

        #endregion Serialization

        #region Static

        public static class Static
        {
            internal static Func<SmartSerializable<T>> GetInstanceFactory { get; set; } = () => new SmartSerializable<T>();

            internal static SmartSerializable<T> GetInstance() => GetInstanceFactory();

            public static T Deserialize(string fileName, string folderPath) =>
                GetInstance().Deserialize(fileName, folderPath);

            public static T Deserialize(string fileName, string folderPath, bool askUserOnError) =>
                GetInstance().Deserialize(fileName, folderPath, askUserOnError);

            public static T Deserialize(string fileName, string folderPath, bool askUserOnError, JsonSerializerSettings settings) =>
                GetInstance().Deserialize(fileName, folderPath, askUserOnError, settings);

            public static T Deserialize<U>(ISmartSerializableLoader<U> loader) where U : class, ISmartSerializable<U>, new() =>
                GetInstance().Deserialize(loader);

            public static T DeseriealizeObject(string json, JsonSerializerSettings settings) =>
                GetInstance().DeserializeObject(json, settings);

            public static async Task<T> DeserializeAsync<U>(ISmartSerializableLoader<U> loader) where U : class, ISmartSerializable<U>, new() =>
                await GetInstance().DeserializeAsync(loader);

            public static async Task<T> DeserializeAsync<U>(ISmartSerializableLoader<U> loader, bool askUserOnError) where U : class, ISmartSerializable<U>, new() =>
                await GetInstance().DeserializeAsync(loader, askUserOnError);

            public static async Task<T> DeserializeAsync<U>(ISmartSerializableLoader<U> loader, bool askUserOnError, Func<T> altLoader) where U : class, ISmartSerializable<U>, new() =>
                await GetInstance().DeserializeAsync(loader, askUserOnError, altLoader);

            internal static JsonSerializerSettings GetDefaultSettings() =>
                SmartSerializable<T>.GetDefaultSettings();
        }

        #endregion Static
    }
}
