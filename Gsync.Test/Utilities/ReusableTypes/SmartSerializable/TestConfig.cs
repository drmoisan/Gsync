using Gsync.Utilities.Interfaces;
using Gsync.Utilities.ReusableTypes;
using Newtonsoft.Json;

namespace Gsync.Test.Utilities.ReusableTypes.SmartSerializable
{    
    public class TestConfig : ISmartSerializable<TestConfig>
    {
        public string Name { get; set; }
        public ISmartSerializableConfig Config { get; set; } = new NewSmartSerializableConfig();

        // The simplest stub for required interface
        public void Serialize() { }
        public void Serialize(string filePath) { }
        public void SerializeThreadSafe(string filePath) { }
        public string SerializeToString() => JsonConvert.SerializeObject(this);
        public void SerializeToStream(System.IO.StreamWriter sw) { }
        public TestConfig Deserialize(string fileName, string folderPath) => new TestConfig();
        public TestConfig Deserialize(string fileName, string folderPath, bool askUserOnError) => new TestConfig();
        public TestConfig Deserialize(string fileName, string folderPath, bool askUserOnError, JsonSerializerSettings settings) => new TestConfig();
        public TestConfig Deserialize<U>(ISmartSerializableLoader<U> loader) where U : class, ISmartSerializable<U>, new() => new TestConfig();
        public TestConfig Deserialize<U>(ISmartSerializableLoader<U> loader, bool askUserOnError, System.Func<TestConfig> altLoader) where U : class, ISmartSerializable<U>, new() => new TestConfig();
        public System.Threading.Tasks.Task<TestConfig> DeserializeAsync<U>(ISmartSerializableLoader<U> loader) where U : class, ISmartSerializable<U>, new() => System.Threading.Tasks.Task.FromResult(new TestConfig());
        public System.Threading.Tasks.Task<TestConfig> DeserializeAsync<U>(ISmartSerializableLoader<U> loader, bool askUserOnError) where U : class, ISmartSerializable<U>, new() => System.Threading.Tasks.Task.FromResult(new TestConfig());
        public System.Threading.Tasks.Task<TestConfig> DeserializeAsync<U>(ISmartSerializableLoader<U> loader, bool askUserOnError, System.Func<TestConfig> altLoader) where U : class, ISmartSerializable<U>, new() => System.Threading.Tasks.Task.FromResult(new TestConfig());
        public TestConfig DeserializeObject(string json, JsonSerializerSettings settings) => JsonConvert.DeserializeObject<TestConfig>(json, settings);
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        public void Notify([System.Runtime.CompilerServices.CallerMemberName] string propertyName = "") => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
    }

}
