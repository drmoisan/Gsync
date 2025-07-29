using Gsync.Utilities.Interfaces;
using Gsync.Utilities.ReusableTypes;
using Newtonsoft.Json;
using System;
using System.ComponentModel;
using System.Threading.Tasks;

namespace Gsync.Test.Utilities.ReusableTypes.SmartSerializable
{
    // Minimal dummy implementation for T
    public class DummySerializable : ISmartSerializable<DummySerializable>
    {
        public string Name { get; set; }
        public NewSmartSerializableConfig Config { get; set; } = new NewSmartSerializableConfig();
        public event PropertyChangedEventHandler PropertyChanged;

        public DummySerializable Deserialize(string fileName, string folderPath) => new DummySerializable();
        public DummySerializable Deserialize(string fileName, string folderPath, bool askUserOnError) => new DummySerializable();
        public DummySerializable Deserialize(string fileName, string folderPath, bool askUserOnError, JsonSerializerSettings settings) => new DummySerializable();
        public DummySerializable Deserialize<U>(SmartSerializable<U> loader) where U : class, ISmartSerializable<U>, new() => new DummySerializable();
        public DummySerializable Deserialize<U>(SmartSerializable<U> loader, bool askUserOnError, Func<DummySerializable> altLoader) where U : class, ISmartSerializable<U>, new() => new DummySerializable();
        public Task<DummySerializable> DeserializeAsync<U>(SmartSerializable<U> config) where U : class, ISmartSerializable<U>, new() => Task.FromResult(new DummySerializable());
        public Task<DummySerializable> DeserializeAsync<U>(SmartSerializable<U> config, bool askUserOnError) where U : class, ISmartSerializable<U>, new() => Task.FromResult(new DummySerializable());
        public Task<DummySerializable> DeserializeAsync<U>(SmartSerializable<U> config, bool askUserOnError, Func<DummySerializable> altLoader) where U : class, ISmartSerializable<U>, new() => Task.FromResult(new DummySerializable());
        public DummySerializable DeserializeObject(string json, JsonSerializerSettings settings) => new DummySerializable();
        public void Serialize() { }
        public void Serialize(string filePath) { }
        public void SerializeThreadSafe(string filePath) { }
    }
}
