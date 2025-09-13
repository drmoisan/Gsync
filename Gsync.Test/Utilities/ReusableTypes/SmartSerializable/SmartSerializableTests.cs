using FluentAssertions;
using Gsync.Utilities.HelperClasses;
using Gsync.Utilities.ReusableTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            _fileSystem = new FakeFileSystem();
            _userDialog = new FakeUserDialog();
            _timerFactory = new FakeTimerFactory();
            _sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = _timerFactory
            };

            SmartSerializable<TestConfig>.Static.GetInstanceFactory = () =>
                new SmartSerializable<TestConfig>(new TestConfig())
                {
                    FileSystem = _fileSystem,
                    UserDialog = _userDialog,
                    TimerFactory = _timerFactory
                };
        }

        [TestMethod]
        public void AskUser_ReturnsDialogResultIfAsking()
        {
            _userDialog.ResultToReturn = DialogResult.No;
            var result = _sut.GetType()
                .GetMethod("AskUser", BindingFlags.NonPublic | BindingFlags.Instance)
                .Invoke(_sut, new object[] { true, "something" });

            result.Should().Be(DialogResult.No);
        }

        [TestMethod]
        public void AskUser_ReturnsYesIfNotAsking()
        {
            var result = _sut.GetType()
                .GetMethod("AskUser", BindingFlags.NonPublic | BindingFlags.Instance)
                .Invoke(_sut, new object[] { false, "msg" });

            result.Should().Be(DialogResult.Yes);
        }

        [TestMethod]
        public void Config_PropertyChanged_TriggersNotify()
        {
            string raised = null;
            _sut.PropertyChanged += (s, e) => raised = e.PropertyName;

            _sut.Config = new NewSmartSerializableConfig();
            _sut.Config.JsonSettings = new JsonSerializerSettings();
            raised.Should().NotBeNull();
        }

        [TestMethod]
        public void Config_PropertyChangedEvent_BubblesUp()
        {
            string raised = null;
            _sut.PropertyChanged += (s, e) => raised = e.PropertyName;
            _sut.Config.ClassifierActivated = true;
            raised.Should().Be(nameof(_sut.Config.ClassifierActivated));
        }

        [TestMethod]
        public void CreateEmpty_ThrowsOnUserNo()
        {
            Action act = () => _sut.GetType()
                .GetMethod("CreateEmpty", BindingFlags.NonPublic | BindingFlags.Instance)
                .Invoke(_sut, new object[] { DialogResult.No, new FilePathHelper("x", "y"), SmartSerializable<TestConfig>.GetDefaultSettings(), null });

            act.Should().Throw<TargetInvocationException>().Where(e => e.InnerException is InvalidOperationException);
            //act.Should().Throw<TargetInvocationException>().Where(e => e.InnerException is ArgumentNullException);
        }

        [TestMethod]
        public void CreateEmpty_UsesAltLoader_WhenProvided()
        {
            bool wasCalled = false;
            Func<TestConfig> altLoader = () => { wasCalled = true; return new TestConfig(); };
            var disk = new FilePathHelper("f.json", "folder");

            var result = typeof(SmartSerializable<TestConfig>)
                .GetMethod("CreateEmpty", BindingFlags.NonPublic | BindingFlags.Instance)
                .Invoke(_sut, new object[] { DialogResult.Yes, disk, SmartSerializable<TestConfig>.GetDefaultSettings(), altLoader });

            wasCalled.Should().BeTrue();
            result.Should().NotBeNull();
        }

        [TestMethod]
        public void CreateStreamWriter_UsesDefaultIfNotInjected()
        {
            _sut.CreateStreamWriter = null; // Should fallback to FileSystem.CreateText
            var sw = _sut.CreateStreamWriter("any.txt");
            sw.Should().NotBeNull();
            sw.Dispose();
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

        [TestMethod]
        public void DeepCopy_ConfigIsPropagated_WhenDeserializedViaLoader()
        {
            var config = new NewSmartSerializableConfig();
            var testConfig = new TestConfig { Name = "Original" };
            var loader = new SmartSerializable<TestConfig>(testConfig)
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = _timerFactory,
                Config = config
            };

            _fileSystem.FileContent = JsonConvert.SerializeObject(testConfig);

            var instance = _sut.Deserialize(loader);

            // This checks deep equality of properties, not object reference
            instance.Config.Should().BeEquivalentTo(loader.Config);
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
            result.Name.Should().Be("z");
        }

        [TestMethod]
        public void Deserialize_CorruptJson_AsksUserAndCreatesNewInstance()
        {
            _fileSystem.FileContent = "not a valid json!";
            _userDialog.ResultToReturn = DialogResult.Yes;
            var result = _sut.Deserialize("corrupt.json", "folder", true);
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

        [TestMethod]
        public void Deserialize_FileNotFound_FileNameFolderNameOverload_ReturnsNewInstance()
        {
            _fileSystem.FileExistsResult = false;
            _userDialog.ResultToReturn = DialogResult.Yes;

            var result = _sut.Deserialize("missing.json", "folder");
            
            result.Should().NotBeNull();
        }

        [TestMethod]
        public void Deserialize_FileNotFound_UserDeclines_Throws()
        {
            _fileSystem.FileExistsResult = false;
            _userDialog.ResultToReturn = DialogResult.No;

            Action act = () => _sut.Deserialize("missing.json", "folder", true);

            act.Should().Throw<InvalidOperationException>();
        }

        [TestMethod]
        public void Deserialize_FileNotFound_UserDialogNull_DoesNotThrow()
        {
            _fileSystem.FileExistsResult = false;
            _sut.UserDialog = null;

            var result = _sut.Deserialize("missing.json", "folder", true);

            result.Should().NotBeNull();
        }

        [TestMethod]
        public void Deserialize_NullLoader_ThrowsArgumentNullException()
        {
            Action act = () => _sut.Deserialize<TestConfig>(null);
            act.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void Deserialize_WithAltLoader_CreatesUsingAltLoader()
        {
            _fileSystem.FileExistsResult = false;
            _userDialog.ResultToReturn = DialogResult.Yes;
            var altWasCalled = false;
            TestConfig AltLoader() { altWasCalled = true; return new TestConfig { Name = "ALT" }; }

            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = _timerFactory
            };

            var result = _sut.Deserialize(loader, true, AltLoader);
            altWasCalled.Should().BeTrue();
            result.Name.Should().Be("ALT");
        }

        [TestMethod]
        public void Deserialize_AltLoaderReturnsNull_Throws()
        {
            _fileSystem.FileExistsResult = false;
            _userDialog.ResultToReturn = DialogResult.Yes;
            Func<TestConfig> altLoader = () => null;

            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = _timerFactory
            };

            Action act = () => _sut.Deserialize(loader, true, altLoader);

            act.Should().Throw<InvalidOperationException>();
        }

        [TestMethod]
        public void DeserializeObject_SetsConfigSettings()
        {
            var json = JsonConvert.SerializeObject(new TestConfig { Name = "foo" });
            var settings = new JsonSerializerSettings();
            var instance = _sut.DeserializeObject(json, settings);

            instance.Config.JsonSettings.Should().NotBeNull();
        }

        [TestMethod]
        public void DeserializeObject_ValidJson_ReturnsInstance()
        {
            var json = JsonConvert.SerializeObject(new TestConfig { Name = "foo" });
            var result = _sut.DeserializeObject(json, SmartSerializable<TestConfig>.GetDefaultSettings());

            result.Should().NotBeNull();
            result.Name.Should().Be("foo");
        }

        [TestMethod]
        public void DeserializeJson_FileDoesNotExist_Throws()
        {
            _fileSystem.FileExistsResult = false;
            var method = typeof(SmartSerializable<TestConfig>).GetMethod("DeserializeJson", BindingFlags.NonPublic | BindingFlags.Instance, null, new Type[] { typeof(FilePathHelper), typeof(JsonSerializerSettings) }, null);

            //var result = method.Invoke(_sut, new object[] { new FilePathHelper("missing.json", ""), SmartSerializable<TestConfig>.GetDefaultSettings() });
            var act = () => method.Invoke(_sut, new object[] { new FilePathHelper("missing.json", ""), SmartSerializable<TestConfig>.GetDefaultSettings() });

            act.Should().Throw<TargetInvocationException>().Where(e => e.InnerException is System.IO.FileNotFoundException);
            //result.Should().BeNull();
        }

        [TestMethod]
        public void DeserializeJson_Throws_CaughtAndLogged()
        {
            _fileSystem.FileExistsResult = true;
            _fileSystem.FileContent = "NOT JSON!";
            var method = typeof(SmartSerializable<TestConfig>).GetMethod("DeserializeJson", BindingFlags.NonPublic | BindingFlags.Instance, null, new Type[] { typeof(FilePathHelper), typeof(JsonSerializerSettings) }, null);

            var result = method.Invoke(_sut, new object[] { new FilePathHelper("corrupt.json", ""), SmartSerializable<TestConfig>.GetDefaultSettings() });
            result.Should().BeNull();
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
        public void PropertyChanged_Handler_DoesNotThrow_WhenNoSubscribers()
        {
            Action act = () => _sut.Notify("abc");
            act.Should().NotThrow();
        }

        [TestMethod]
        public void ParentNull_ThrowsOnSerializeThreadSafe()
        {
            var sut = new SmartSerializable<TestConfig>(null)
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = _timerFactory
            };
            sut.Config.Disk.FilePath = "test.json";
            Action act = () => sut.SerializeThreadSafe("test.json");
            act.Should().Throw<Exception>()
                .Where(e => e is ArgumentNullException || e is NullReferenceException);
        }


        [TestMethod]
        public void RequestSerialization_CalledTwice_OnlyOneTimerFires()
        {
            var fakeTimerFactory = _sut.TimerFactory as FakeTimerFactory;
            fakeTimerFactory.ImmediateTimer = false;
            _sut.Config.Disk.FilePath = "multi.json";

            _sut.Serialize();
            var firstTimer = fakeTimerFactory.LastTimer;

            _sut.Serialize(); // Should not create a second timer
            var secondTimer = fakeTimerFactory.LastTimer;

            firstTimer.Should().BeSameAs(secondTimer);
        }

        [TestMethod]
        public void RequestSerialization_OnlyCreatesOneTimer_WhenCalledRepeatedly()
        {
            var fakeFactory = new FakeTimerFactory();
            _sut.TimerFactory = fakeFactory;
            _sut.Config.Disk.FilePath = "f.json";
            _sut.Serialize();
            _sut.Serialize();

            fakeFactory.LastTimer.Should().NotBeNull();
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
        public void Serialize_DoesNothing_WhenFilePathEmpty()
        {
            _sut.Config.Disk.FilePath = string.Empty;
            Action act = () => _sut.Serialize();
            act.Should().NotThrow();
            _fileSystem.WrittenPath.Should().BeNullOrEmpty();
        }

        [TestMethod]
        public void Serialize_EmptyFilePath_DoesNothing()
        {
            _sut.Config.Disk.FilePath = "";
            _sut.Serialize();

            _fileSystem.WrittenPath.Should().BeNull();
        }

        [TestMethod]
        public void SerializeToString_ReturnsValidJson()
        {
            var json = _sut.SerializeToString();
            json.Should().Contain("{");
        }

        [TestMethod]
        public void SerializeToString_UsesCustomJsonSettings()
        {
            _sut.Config.JsonSettings = new JsonSerializerSettings { Formatting = Formatting.None };
            var json = _sut.SerializeToString();
            json.Should().NotContain(Environment.NewLine);
        }

        [TestMethod]
        public void SerializeThreadSafe_ConcurrentCalls_AreThreadSafe()
        {
            _sut.Config.Disk.FilePath = "concurrent.json";
            var tasks = new[]
            {
                Task.Run(() => _sut.SerializeThreadSafe("concurrent.json")),
                Task.Run(() => _sut.SerializeThreadSafe("concurrent.json"))
            };
            Task.WaitAll(tasks);
            _fileSystem.WrittenPath.Should().Be("concurrent.json");
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
        public void Static_GetDefaultSettings_ReturnsValidSettings()
        {
            var settings = SmartSerializable<TestConfig>.Static.GetDefaultSettings();
            settings.Should().NotBeNull();
            settings.TypeNameHandling.Should().Be(TypeNameHandling.Auto);
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
        public void Name_CanSetAndGet()
        {
            _sut.Name = "SampleName";
            _sut.Name.Should().Be("SampleName");
        }

        [TestMethod]
        public void Config_CanBeSetAndGet()
        {
            var config = new NewSmartSerializableConfig();
            _sut.Config = config;
            _sut.Config.Should().BeSameAs(config);
        }

        [TestMethod]
        public void CreateStreamWriter_DefaultsToFileSystemCreateText()
        {
            _sut.CreateStreamWriter = null; // Should default to FileSystem.CreateText
            var sw = _sut.CreateStreamWriter("default.txt");
            sw.Should().NotBeNull();
            sw.Dispose();
        }

        [TestMethod]
        public void CreateStreamWriter_CanBeOverridden()
        {
            bool called = false;
            _sut.CreateStreamWriter = path => { called = true; return new StreamWriter(new MemoryStream()); };
            using (var sw = _sut.CreateStreamWriter("foo.txt")) { }
            called.Should().BeTrue();
        }

        [TestMethod]
        public void PropertyChanged_CanBeSubscribedAndUnsubscribed()
        {
            PropertyChangedEventHandler handler = (s, e) => { };
            _sut.PropertyChanged += handler;
            _sut.PropertyChanged -= handler;
            // No assertion needed: just verifies no exception is thrown
        }

        [TestMethod]
        public void Constructor_WithoutParent_InitializesConfig()
        {
            var ss = new SmartSerializable<TestConfig>();
            ss.Config.Should().NotBeNull();
        }

        [TestMethod]
        public void FileSystem_CanBeSetDynamically()
        {
            var fs = new FakeFileSystem();
            _sut.FileSystem = fs;
            _sut.FileSystem.Should().BeSameAs(fs);
        }

        [TestMethod]
        public void UserDialog_CanBeSetDynamically()
        {
            var ud = new FakeUserDialog();
            _sut.UserDialog = ud;
            _sut.UserDialog.Should().BeSameAs(ud);
        }

        [TestMethod]
        public void TimerFactory_CanBeSetDynamically()
        {
            var tf = new FakeTimerFactory();
            _sut.TimerFactory = tf;
            _sut.TimerFactory.Should().BeSameAs(tf);
        }

        [TestMethod]
        public void Static_Deserialize_Overloads_Work()
        {
            var fileName = "test.json";
            var folder = "f";
            var json = JsonConvert.SerializeObject(new TestConfig { Name = "x" });
            File.WriteAllText(fileName, json);
            try
            {
                var inst = SmartSerializable<TestConfig>.Static.Deserialize(fileName, folder);
                inst.Should().NotBeNull();
            }
            finally
            {
                if (File.Exists(fileName)) File.Delete(fileName);
            }
        }

        [TestMethod]
        public void Static_Deserialize_WithAskUserOnError_Works()
        {
            var fileName = "test2.json";
            var folder = "f";
            File.WriteAllText(fileName, JsonConvert.SerializeObject(new TestConfig { Name = "y" }));
            try
            {
                var inst = SmartSerializable<TestConfig>.Static.Deserialize(fileName, folder, true);
                inst.Should().NotBeNull();
            }
            finally
            {
                if (File.Exists(fileName)) File.Delete(fileName);
            }
        }

        [TestMethod]
        public void Static_Deserialize_WithSettings_Works()
        {
            var fileName = "test3.json";
            var folder = "f";
            File.WriteAllText(fileName, JsonConvert.SerializeObject(new TestConfig { Name = "z" }));
            try
            {
                var settings = SmartSerializable<TestConfig>.GetDefaultSettings();
                var inst = SmartSerializable<TestConfig>.Static.Deserialize(fileName, folder, true, settings);
                inst.Should().NotBeNull();
            }
            finally
            {
                if (File.Exists(fileName)) File.Delete(fileName);
            }
        }

        [TestMethod]
        public async Task Static_DeserializeAsync_Works()
        {
            var expectedObject = new TestConfig { Name = "test object" };
            var json = JsonConvert.SerializeObject(expectedObject);
            _fileSystem.FileContent = json;
            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = _timerFactory
            };
            var actualObject = await SmartSerializable<TestConfig>.Static.DeserializeAsync(loader);
            actualObject.Should().NotBeNull();
            actualObject.Should().BeEquivalentTo(expectedObject);

        }


        [TestMethod]
        public void Static_DeseriealizeObject_Works()
        {
            var json = JsonConvert.SerializeObject(new TestConfig { Name = "a" });
            var inst = SmartSerializable<TestConfig>.Static.DeseriealizeObject(json, SmartSerializable<TestConfig>.GetDefaultSettings());
            inst.Should().NotBeNull();
            inst.Name.Should().Be("a");
        }

        [TestMethod]
        public void Static_GetDefaultSettings_Works()
        {
            var s = SmartSerializable<TestConfig>.Static.GetDefaultSettings();
            s.Should().NotBeNull();
        }

        [TestMethod]
        public void DeserializeJson_WithValidJson_ReturnsDeserializedInstance()
        {
            // Arrange
            var testConfig = new TestConfig { Name = "Hello" };
            var fileSystem = new FakeFileSystem
            {
                FileExistsResult = true,
                FileContent = JsonConvert.SerializeObject(testConfig)
            };
            var testSut = new TestableSmartSerializable(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = new FakeUserDialog(),
                TimerFactory = new FakeTimerFactory()
            };

            var disk = new FilePathHelper("somefile.json", "");

            // Act
            var result = testSut.CallDeserializeJson(disk);

            // Assert
            result.Should().NotBeNull();
            result.Name.Should().Be("Hello");
        }

        [TestMethod]
        public void DeserializeJson_FileDoesNotExist_Throws2()
        {
            // Arrange
            //var fileSystem = new FakeFileSystem
            //{
            //    FileExistsResult = false
            //};
            _fileSystem.FileExistsResult = false; // Use the same file system as in the test class
            var testSut = new TestableSmartSerializable(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = new FakeUserDialog(),
                TimerFactory = new FakeTimerFactory()
            };

            var disk = new FilePathHelper("missing.json", "");
            
            // Act
            var act = () => testSut.CallDeserializeJson(disk);
            // Assert
            act.Should().Throw<System.IO.FileNotFoundException>();
            //var result = testSut.CallDeserializeJson(disk);

            //// Assert
            //result.Should().BeNull();
        }

        [TestMethod]
        public async Task DeserializeAsync_WithAskUserOnError_True_FileExists_ReturnsDeserializedInstance()
        {
            // Arrange
            var expected = new TestConfig { Name = "Async1" };
            var fileSystem = new FakeFileSystem
            {
                FileExistsResult = true,
                FileContent = JsonConvert.SerializeObject(expected)
            };

            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = new FakeUserDialog(),
                TimerFactory = new FakeTimerFactory()
            };

            var sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = new FakeUserDialog(),
                TimerFactory = new FakeTimerFactory()
            };

            // Act
            var actual = await sut.DeserializeAsync(loader, askUserOnError: true);

            // Assert
            actual.Should().NotBeNull();
            actual.Name.Should().Be("Async1");
        }

        [TestMethod]
        public async Task DeserializeAsync_WithAskUserOnError_True_FileMissing_CreatesNewInstance()
        {
            // Arrange
            var fileSystem = new FakeFileSystem { FileExistsResult = false };
            var userDialog = new FakeUserDialog { ResultToReturn = DialogResult.Yes };

            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = userDialog,
                TimerFactory = new FakeTimerFactory()
            };

            var sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = userDialog,
                TimerFactory = new FakeTimerFactory()
            };

            // Act
            var actual = await sut.DeserializeAsync(loader, askUserOnError: true);

            // Assert
            actual.Should().NotBeNull();
        }

        [TestMethod]
        public async Task DeserializeAsync_WithAltLoader_FileMissing_UsesAltLoader()
        {
            // Arrange
            var fileSystem = new FakeFileSystem { FileExistsResult = false };
            var userDialog = new FakeUserDialog { ResultToReturn = DialogResult.Yes };
            var wasAltLoaderCalled = false;

            Func<TestConfig> altLoader = () => { wasAltLoaderCalled = true; return new TestConfig { Name = "ALT_ASYNC" }; };

            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = userDialog,
                TimerFactory = new FakeTimerFactory()
            };

            var sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = userDialog,
                TimerFactory = new FakeTimerFactory()
            };

            // Act
            var actual = await sut.DeserializeAsync(loader, askUserOnError: true, altLoader);

            // Assert
            actual.Should().NotBeNull();
            actual.Name.Should().Be("ALT_ASYNC");
            wasAltLoaderCalled.Should().BeTrue();
        }

        [TestMethod]
        public async Task DeserializeAsync_WithAltLoader_UserDeclines_Throws()
        {
            // Arrange
            var fileSystem = new FakeFileSystem { FileExistsResult = false };
            var userDialog = new FakeUserDialog { ResultToReturn = DialogResult.No };

            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = userDialog,
                TimerFactory = new FakeTimerFactory()
            };

            var sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = userDialog,
                TimerFactory = new FakeTimerFactory()
            };

            Func<TestConfig> altLoader = () => new TestConfig();

            // Act
            Func<Task> act = async () => await sut.DeserializeAsync(loader, askUserOnError: true, altLoader);

            // Assert
            await act.Should().ThrowAsync<InvalidOperationException>();
        }

        [TestMethod]
        public async Task Static_DeserializeAsync_WithAskUserOnError_FileExists_ReturnsInstance()
        {
            // Arrange
            var expected = new TestConfig { Name = "AsyncStatic1" };
            _fileSystem.FileExistsResult = true; // Use the same file system as in the test class
            _fileSystem.FileContent = JsonConvert.SerializeObject(expected);
            

            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = new FakeTimerFactory()
            };

            // Act
            var actual = await SmartSerializable<TestConfig>.Static.DeserializeAsync(loader, askUserOnError: true);

            // Assert
            actual.Should().NotBeNull();
            actual.Name.Should().Be("AsyncStatic1");
        }

        [TestMethod]
        public async Task Static_DeserializeAsync_WithAltLoader_FileMissing_UsesAltLoader()
        {
            // Arrange            
            _fileSystem.FileExistsResult = false; // Use the same file system as in the test class
            var userDialog = new FakeUserDialog { ResultToReturn = DialogResult.Yes };
            var wasAltLoaderCalled = false;

            Func<TestConfig> altLoader = () => { wasAltLoaderCalled = true; return new TestConfig { Name = "ALT_STATIC_ASYNC" }; };

            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = userDialog,
                TimerFactory = new FakeTimerFactory()
            };

            // Act
            var actual = await SmartSerializable<TestConfig>.Static.DeserializeAsync(loader, askUserOnError: true, altLoader);

            // Assert
            actual.Should().NotBeNull();
            actual.Name.Should().Be("ALT_STATIC_ASYNC");
            wasAltLoaderCalled.Should().BeTrue();
        }

        [TestMethod]
        public async Task Static_DeserializeAsync_WithAltLoader_UserDeclines_Throws()
        {
            // Arrange
            _fileSystem.FileExistsResult = false; // Use the same file system as in the test class
            //var userDialog = new FakeUserDialog { ResultToReturn = DialogResult.No };
            _userDialog.ResultToReturn = DialogResult.No; // Use the same user dialog as in the test class

            var loader = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = _fileSystem,
                UserDialog = _userDialog,
                TimerFactory = new FakeTimerFactory()
            };

            Func<TestConfig> altLoader = () => new TestConfig();

            // Act
            Func<Task> act = async () => await SmartSerializable<TestConfig>.Static.DeserializeAsync(loader, askUserOnError: true, altLoader);

            // Assert
            await act.Should().ThrowAsync<InvalidOperationException>();
        }

        [TestMethod]
        public void Serialize_WithFilePath_SetsConfigDiskFilePath()
        {
            // Arrange
            var path = "somepath.json";
            
            // Swap out the RequestSerialization method with a delegate to capture the call.
            var sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = new FakeFileSystem(),
                UserDialog = new FakeUserDialog(),
                TimerFactory = new FakeTimerFactory()
            };
            // Use reflection to override the protected RequestSerialization (if you want to check that it's called)
            var requestSerializationMethod = typeof(SmartSerializable<TestConfig>).GetMethod("RequestSerialization", BindingFlags.NonPublic | BindingFlags.Instance);

            // Act
            sut.Serialize(path);

            // Assert
            sut.Config.Disk.FilePath.Should().Be(path);
            // Optionally, you could verify the file was actually written, if dependencies are set.
        }

        [TestMethod]
        public void Serialize_WithFilePath_TriggersSerialization()
        {
            // Arrange
            var path = "serialize.json";
            var fileSystem = new FakeFileSystem();
            var timerFactory = new FakeTimerFactory { ImmediateTimer = true }; // Ensures immediate fire
            var sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = fileSystem,
                UserDialog = new FakeUserDialog(),
                TimerFactory = timerFactory
            };

            // Act
            sut.Serialize(path);

            // Assert
            fileSystem.WrittenPath.Should().Be(path);
        }

        [TestMethod]
        public void Serialize_WithEmptyFilePath_SetsConfigDiskFilePathButDoesNotThrow()
        {
            // Arrange
            var path = string.Empty;
            var sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = new FakeFileSystem(),
                UserDialog = new FakeUserDialog(),
                TimerFactory = new FakeTimerFactory()
            };

            // Act
            Action act = () => sut.Serialize(path);

            // Assert
            act.Should().NotThrow();
            sut.Config.Disk.FilePath.Should().Be(path);
            // No file should be written
            (sut.FileSystem as FakeFileSystem).WrittenPath.Should().BeNullOrEmpty();
        }

        [TestMethod]
        public void DeserializeObject_InvalidJson_CatchesExceptionAndReturnsNull()
        {
            // Arrange
            var sut = new SmartSerializable<TestConfig>(new TestConfig())
            {
                FileSystem = new FakeFileSystem(),
                UserDialog = new FakeUserDialog(),
                TimerFactory = new FakeTimerFactory()
            };
            var invalidJson = "{ not valid json!"; // Deliberately invalid

            // Act
            Action act = () => sut.DeserializeObject(invalidJson, SmartSerializable<TestConfig>.GetDefaultSettings());
            var result = sut.DeserializeObject(invalidJson, SmartSerializable<TestConfig>.GetDefaultSettings());

            // Assert
            act.Should().NotThrow(); // Method handles the exception
            result.Should().BeNull();
        }



        [TestCleanup]
        public void Cleanup()
        {
            SmartSerializable<TestConfig>.Static.GetInstanceFactory = () => new SmartSerializable<TestConfig>();
        }


    }
}
