using Newtonsoft.Json;
using System;
using System.ComponentModel;
using System.Threading.Tasks;
using Gsync.Utilities.ReusableTypes;

namespace Gsync.Utilities.Interfaces
{
    public interface ISmartSerializable<T>:INotifyPropertyChanged where T: class, ISmartSerializable<T>, new()
    {
        T Deserialize(string fileName, string folderPath);
        T Deserialize(string fileName, string folderPath, bool askUserOnError);
        T Deserialize(string fileName, string folderPath, bool askUserOnError, JsonSerializerSettings settings);
        T Deserialize<U>(SmartSerializable<U> loader) where U : class, ISmartSerializable<U>, new();
        T Deserialize<U>(SmartSerializable<U> loader, bool askUserOnError, Func<T> altLoader)
            where U : class, ISmartSerializable<U>, new();
        Task<T> DeserializeAsync<U>(SmartSerializable<U> config) where U : class, ISmartSerializable<U>, new();
        Task<T> DeserializeAsync<U>(SmartSerializable<U> config, bool askUserOnError) where U : class, ISmartSerializable<U>, new();
        Task<T> DeserializeAsync<U>(SmartSerializable<U> config, bool askUserOnError, Func<T> altLoader) where U : class, ISmartSerializable<U>, new();
        T DeserializeObject(string json, JsonSerializerSettings settings);

        void Serialize();
        void Serialize(string filePath);
        void SerializeThreadSafe(string filePath);
        

        NewSmartSerializableConfig Config { get; set; }

        string Name { get; set; }

    }
}