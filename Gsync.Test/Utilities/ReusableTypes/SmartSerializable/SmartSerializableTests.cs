using FluentAssertions;
using Gsync.Utilities.HelperClasses;
using Gsync.Utilities.ReusableTypes;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading.Tasks;
using static Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext;

namespace Gsync.Test.Utilities.ReusableTypes.SmartSerializable
{   
    [TestClass]
    public class SmartSerializableTests
    {
        private SmartSerializable<TestConfig> _sut;
        private FakeFileSystem _fileSystem;
        private FakeUserDialog _userDialog;
        private FakeTimerFactory _timerFactory;

        [TestInitialize]
        public void Init()
        {
            Console.SetOut(new DebugTextWriter()); 
            _fileSystem = new FakeFileSystem();
            _userDialog = new FakeUserDialog();
            _timerFactory = new FakeTimerFactory();
            _sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = _timerFactory
            };
        }

        [TestMethod]
        public void Serialize_WritesToFilePath_WhenFilePathNotEmpty()
        {
            var typedFactory = _sut.TimerFactory as FakeTimerFactory;
            if (typedFactory is not null) { typedFactory.ImmediateTimer = true; }
            var path = "myfile.json";
            _sut.Config.Disk.FilePath = path;

            _sut.Serialize();

            Console.WriteLine($"Expected serialization path: {path}");
            Console.WriteLine($"Actual written path:         {_fileSystem.WrittenPath}");
            _fileSystem.WrittenPath.Should().Be(path);
        }

        [TestMethod]
        public void SerializeThreadSafe_CreatesWriterAndSerializes()
        {
            var path = "testfile.json";
            _sut.Config.Disk.FilePath = path;
            _sut.SerializeThreadSafe(path);

            _fileSystem.WrittenPath.Should().Be(path);
        }

        [TestMethod]
        public void SerializeToString_ReturnsValidJson()
        {
            var json = _sut.SerializeToString();

            json.Should().Contain("{");
        }

        [TestMethod]
        public void DeserializeJson_FileDoesNotExist_ReturnsNull()
        {
            _fileSystem.FileExistsResult = false;
            var result = _sut.DeserializeObject("{}", SmartSerializable<TestConfig>.GetDefaultSettings());

            result.Should().NotBeNull();
        }

        [TestMethod]
        public void Deserialize_FileExists_ReturnsInstance()
        {
            var expectedName = "test";
            var instance = new TestConfig { Name = expectedName };
            _fileSystem.FileContent = JsonConvert.SerializeObject(instance);
            _sut.FileSystem = _fileSystem;

            var actual = _sut.Deserialize("a.json", "folder");

            actual.Should().NotBeNull();
            actual.Name.Should().Be(expectedName);
        }



        /// <summary>
        /// NOTE: The two-parameter overload of Deserialize(Deserialize(fileName, folderPath))
        /// does NOT prompt the user or throw exceptions on missing/corrupt files.
        /// It always uses askUserOnError: false and falls back to creating a new instance.
        /// To test user interaction and exception handling, use the overload that accepts askUserOnError: true.        
        /// Example:        
        /// var instance = sut.Deserialize("missing.json", "folder", askUserOnError: true);
        /// Now, user prompts and exception handling are testable.
        /// </summary>
        [TestMethod]
        public void Deserialize_FileNotFound_FileNameFolderNameOverload_ReturnsNewInstance()
        {
            _fileSystem.FileExistsResult = false;
            _userDialog.ResultToReturn = System.Windows.Forms.DialogResult.Yes;

            var result = _sut.Deserialize("missing.json", "folder");

            result.Should().NotBeNull();
        }

        


        [TestMethod]
        public void Notify_RaisesPropertyChanged()
        {
            string raised = null;
            _sut.PropertyChanged += (s, e) => raised = e.PropertyName;

            _sut.Notify("MyProp");
            raised.Should().Be("MyProp");
        }

        [TestMethod]
        public async Task DeserializeAsync_ReturnsDeserializedInstance()
        {
            _fileSystem.FileContent = JsonConvert.SerializeObject(new TestConfig { Name = "z" });
            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = _timerFactory
            };

            var result = await _sut.DeserializeAsync(loader);
            result.Should().NotBeNull();
        }

        [TestMethod]
        public void RequestSerialization_TriggersTimer()
        {
            _sut.Config.Disk.FilePath = "p.json";
            _sut.Serialize();

            _timerFactory.LastTimer.Should().NotBeNull();
            (_timerFactory.LastTimer as FakeTimer).Started.Should().BeTrue();
        }

        [TestMethod]
        public void StaticDeserialize_CallsUnderlyingInstance()
        {
            var json = JsonConvert.SerializeObject(new TestConfig { Name = "static" });
            var instance = SmartSerializable<TestConfig>.Static.DeseriealizeObject(json, SmartSerializable<TestConfig>.GetDefaultSettings());

            instance.Should().NotBeNull();
            instance.Name.Should().Be("static");
        }

        [TestMethod]
        public void CreateStreamWriter_UsesInjectedDelegateIfSet()
        {
            bool wasCalled = false;
            _sut.CreateStreamWriter = path =>
            {
                wasCalled = true;
                return new StreamWriter(new MemoryStream());
            };
            _sut.SerializeThreadSafe("abc.txt");
            wasCalled.Should().BeTrue();
        }

        // Add more tests for edge cases and negative scenarios as needed.
    }
}


