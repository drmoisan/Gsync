using Newtonsoft.Json;
using System;
using System.Threading.Tasks;
using Gsync.Utilities.ReusableTypes;

namespace Gsync.Utilities.Interfaces
{
    public interface ISmartSerializableNonTyped
    {
        T Deserialize<T, U>(SmartSerializable<U> config)
            where T : class, ISmartSerializable<T>, new()
            where U : class, ISmartSerializable<U>, new();
        T Deserialize<T>(string fileName, string folderPath) where T : class, ISmartSerializable<T>, new();
        T Deserialize<T>(string fileName, string folderPath, bool askUserOnError) where T : class, ISmartSerializable<T>, new();
        T Deserialize<T>(string fileName, string folderPath, bool askUserOnError, JsonSerializerSettings settings) where T : class, ISmartSerializable<T>, new();
        Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config)
            where T : class, new()
            where U : class, ISmartSerializable<U>, new();
        Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config, bool askUserOnError)
            where T : class, new()
            where U : class, ISmartSerializable<U>, new();
        Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config, bool askUserOnError, Func<T> altLoader)
            where T : class, new()
            where U : class, ISmartSerializable<U>, new();
        T DeserializeObject<T>(string json, JsonSerializerSettings settings) where T : class;
        bool IsSmartSerializable(Type type);
        bool IsSmartSerializable<T>(T instance);
    }
}