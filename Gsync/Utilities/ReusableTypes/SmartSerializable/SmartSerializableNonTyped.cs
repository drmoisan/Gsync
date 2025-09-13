using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Gsync.Utilities.Interfaces;
using Gsync.Utilities.Extensions;


namespace Gsync.Utilities.ReusableTypes
{
    public class SmartSerializableNonTyped() : ISmartSerializableNonTyped
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public bool IsSmartSerializable<T>(T instance)
        {
            return IsSmartSerializable(instance.GetType());
        }

        public bool IsSmartSerializable(Type type)
        {
            return type.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ISmartSerializable<>));
        }

        private SmartSerializableBase GetInstance() => new();

        private SmartSerializable<T> GetInstance<T>()
            where T : class, ISmartSerializable<T>, new() => new();

        public T Deserialize<T>(string fileName, string folderPath)
            where T : class, ISmartSerializable<T>, new() =>
            GetInstance<T>().Deserialize(fileName, folderPath);

        public T Deserialize<T>(string fileName, string folderPath, bool askUserOnError)
            where T : class, ISmartSerializable<T>, new() =>
            GetInstance<T>().Deserialize(fileName, folderPath, askUserOnError);

        public T Deserialize<T>(string fileName, string folderPath, bool askUserOnError, JsonSerializerSettings settings)
            where T : class, ISmartSerializable<T>, new() =>
            GetInstance<T>().Deserialize(fileName, folderPath, askUserOnError, settings);

        public T Deserialize<T, U>(SmartSerializable<U> config)
            where T : class, ISmartSerializable<T>, new()
            where U : class, ISmartSerializable<U>, new() =>
            GetInstance<T>().Deserialize(config);

        public T DeserializeObject<T>(string json, JsonSerializerSettings settings) where T : class
        {
            T instance = default;
            try
            {
                instance = JsonConvert.DeserializeObject<T>(json, settings);
            }
            catch (Exception e)
            {
                logger.Error(e.Message, e);
            }
            if (instance is not null && IsSmartSerializable(instance))
            {
                var config = typeof(T).GetProperty("Config").GetValue(instance) as NewSmartSerializableConfig;
                if (config is not null)
                {
                    config.JsonSettings = settings.DeepCopy();
                }
            }

            return instance;
        }

        public async Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config)
            where T : class, new()
            where U : class, ISmartSerializable<U>, new() =>
            await GetInstance().DeserializeAsync<T, U>(config);

        public async Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config, bool askUserOnError)
            where T : class, new()
            where U : class, ISmartSerializable<U>, new() =>
            await GetInstance().DeserializeAsync<T, U>(config, askUserOnError);

        public async Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config, bool askUserOnError, Func<T> altLoader)
            where T : class, new()
            where U : class, ISmartSerializable<U>, new() =>
            await GetInstance().DeserializeAsync<T, U>(config, askUserOnError, altLoader);



    }
}
