using System;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using Gsync.Utilities.ReusableTypes;
using Gsync.Utilities.HelperClasses;
using Gsync.Utilities.HelperClasses.NewtonSoft;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Newtonsoft.Json;
using Gsync.Utilities.Interfaces;

namespace Gsync.Test.Utilities.ReusableTypes.SmartSerializable
{
    [TestClass]
    public class SmartSerializableTests
    {
        [TestMethod]
        public void Constructor_WithParent_SetsParentAndConfig()
        {
            var dummy = new DummySerializable();
            var smart = new SmartSerializable<DummySerializable>(dummy);

            // Use reflection to access protected _parent
            var parentField = typeof(SmartSerializable<DummySerializable>)
                .GetField("_parent", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var parentValue = parentField.GetValue(smart);

            Assert.AreEqual(dummy, parentValue);
            Assert.IsNotNull(smart.Config);
        }

        [TestMethod]
        public void Config_Setter_SetsPropertySilently()
        {
            var smart = new SmartSerializable<DummySerializable>();
            bool raised = false;
            smart.PropertyChanged += (s, e) => { if (e.PropertyName == "Config") raised = true; };

            var expected = !smart.Config.ClassifierActivated;
            //var mockConfig = new Mock<NewSmartSerializableConfig>();
            //mockConfig.Setup(c => c.ClassifierActivated).Returns(expected);
            //var newConfig = mockConfig.Object;
            var newConfig = new NewSmartSerializableConfig();
            newConfig.ClassifierActivated = expected;
            smart.Config = newConfig;
            bool actual = smart.Config.ClassifierActivated;
            Assert.IsFalse(raised);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void Notify_RaisesPropertyChanged()
        {
            var smart = new SmartSerializable<DummySerializable>();
            bool raised = false;
            smart.PropertyChanged += (s, e) => { if (e.PropertyName == "TestProp") raised = true; };

            smart.Notify("TestProp");

            Assert.IsTrue(raised);
        }

        [TestMethod]
        public void SerializeToString_ProducesValidJson()
        {
            var dummy = new DummySerializable { Name = "Test" };
            var smart = new SmartSerializable<DummySerializable>(dummy);

            string json = smart.SerializeToString();

            Assert.IsTrue(json.Contains("Test"));
        }

        [TestMethod]
        public void DeserializeObject_ReturnsObject()
        {
            var smart = new SmartSerializable<DummySerializable>();
            string json = JsonConvert.SerializeObject(new DummySerializable { Name = "Test" });

            var result = smart.DeserializeObject(json, SmartSerializable<DummySerializable>.GetDefaultSettings());

            Assert.IsNotNull(result);
            Assert.AreEqual("Test", result.Name);
        }

        [TestMethod]
        public async Task DeserializeAsync_ReturnsObject()
        {
            var smart = new SmartSerializable<DummySerializable>();
            var loader = new SmartSerializable<DummySerializable>(new DummySerializable { Name = "AsyncTest" });

            var result = await smart.DeserializeAsync(loader);

            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void SerializeThreadSafe_ThrowsIfParentNull()
        {
            var smart = new SmartSerializable<DummySerializable>();
            // Set up a file path that won't be used
            Assert.ThrowsException<ArgumentNullException>(() => smart.SerializeThreadSafe("dummy.json"));
        }

        [TestMethod]
        public void Config_PropertyChanged_RaisesPropertyChangedWithCorrectName()
        {
            // Arrange
            var smart = new SmartSerializable<DummySerializable>();
            string receivedPropertyName = null;
            smart.PropertyChanged += (s, e) => receivedPropertyName = e.PropertyName;

            // Act: change a property on Config to trigger PropertyChanged
            var oldValue = smart.Config.ClassifierActivated;
            smart.Config.ClassifierActivated = !oldValue;

            // Assert
            Assert.AreEqual(nameof(smart.Config.ClassifierActivated), receivedPropertyName);
        }
    }
}