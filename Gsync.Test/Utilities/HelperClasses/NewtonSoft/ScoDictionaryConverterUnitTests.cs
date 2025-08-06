using FluentAssertions;
using Gsync.Utilities.HelperClasses;
using Gsync.Utilities.HelperClasses.NewtonSoft;
using Gsync.Utilities.ReusableTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using static Gsync.Test.Utilities.HelperClasses.NewtonSoft.ScoDictionaryConverterIntegrationTests;

namespace Gsync.Test.Utilities.HelperClasses.NewtonSoft
{
    [TestClass]
    public class ScoDictionaryConverterUnitTests
    {
        // Simple type to test with
        public class DummyDict : ScoDictionaryNew<string, int> { }

        [TestMethod]
        public void CanConvert_ShouldReturnTrue_ForScoDictionaryNewDerived()
        {
            var converter = new ScoDictionaryConverter();
            var result = converter.CanConvert(typeof(DummyDict));
            result.Should().BeTrue();
        }

        [TestMethod]
        public void CanConvert_ShouldReturnFalse_ForNonDictionaryType()
        {
            var converter = new ScoDictionaryConverter();
            var result = converter.CanConvert(typeof(string));
            result.Should().BeFalse();
        }

        [TestMethod]
        public void WriteJson_ThrowsArgumentNullException_WhenWriterIsNull()
        {
            var converter = new ScoDictionaryConverter();
            Action act = () => converter.WriteJson(null, new DummyDict(), new JsonSerializer());
            act.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void WriteJson_ThrowsArgumentNullException_WhenValueIsNull()
        {
            var converter = new ScoDictionaryConverter();
            Action act = () => converter.WriteJson(new JTokenWriter(), null, new JsonSerializer());
            act.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void WriteJson_ThrowsArgumentNullException_WhenSerializerIsNull()
        {
            var converter = new ScoDictionaryConverter();
            Action act = () => converter.WriteJson(new JTokenWriter(), new DummyDict(), null);
            act.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void ReadJson_ThrowsArgumentNullException_WhenReaderIsNull()
        {
            var converter = new ScoDictionaryConverter();
            Action act = () => converter.ReadJson(null, typeof(DummyDict), null, new JsonSerializer());
            act.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void ReadJson_ThrowsArgumentNullException_WhenObjectTypeIsNull()
        {
            var converter = new ScoDictionaryConverter();
            var reader = new JTokenReader(new JObject());
            Action act = () => converter.ReadJson(reader, null, null, new JsonSerializer());
            act.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void ReadJson_ThrowsArgumentNullException_WhenSerializerIsNull()
        {
            var converter = new ScoDictionaryConverter();
            var reader = new JTokenReader(new JObject());

            Action act = () => converter.ReadJson(reader, typeof(DummyDict), null, null);

            act.Should().Throw<TargetInvocationException>()
               .WithInnerException<ArgumentNullException>()
               .WithMessage("*RemainingObject*");
        }


        [TestMethod]
        public void ReadJson_ReturnsExpectedInstance_WhenValidInput()
        {
            // Arrange: Use a known good json for DummyDict
            var converter = new ScoDictionaryConverter();
            var dummy = new DummyDict();
            dummy.TryAdd("test", 123);

            // Serialize using the integration method
            var settings = new JsonSerializerSettings();
            settings.Converters.Add(converter);
            var json = JsonConvert.SerializeObject(dummy, settings);

            // Act
            var deserialized = JsonConvert.DeserializeObject<DummyDict>(json, settings);

            // Assert
            deserialized.Should().NotBeNull();
            deserialized.Should().BeEquivalentTo(dummy);
        }

        [TestMethod]
        public void WriteJson_SerializesDictionary_AsExpected()
        {
            var converter = new ScoDictionaryConverter();
            var dummy = new DummyDict();
            dummy.TryAdd("a", 1);

            var writer = new JTokenWriter();
            var serializer = new JsonSerializer();
            serializer.Converters.Add(converter);

            converter.WriteJson(writer, dummy, serializer);

            var token = writer.Token as JObject;
            token.Should().NotBeNull();
            token["CoDictionary"].Should().NotBeNull();
        }

        [TestMethod]
        public void ReadJson_ManualCopy_WhenTypeIsNotAssignable_CopiesCommonFields()
        {
            // Arrange
            var derived = new DerivedForEdgeTest
            {
                AdditionalField1 = "present-in-derived",
                Name = "EdgeCase"
            };
            derived.TryAdd("key1", 100);

            // Serialize as derived type
            var json = JsonConvert.SerializeObject(derived, new ScoDictionaryConverter<DerivedForEdgeTest, string, int>());

            // Deserialize with base type (simulate type mismatch)
            var converter = new ScoDictionaryConverter();
            var settings = new JsonSerializerSettings();
            settings.Converters.Add(converter);

            // Act
            var result = JsonConvert.DeserializeObject<ScoDictionaryNew<string, int>>(json, settings);

            // Assert: Should be base type, not derived
            result.Should().BeOfType<ScoDictionaryNew<string, int>>();
            result.ContainsKey("key1").Should().BeTrue();
            result["key1"].Should().Be(100);
            result.Name.Should().Be("EdgeCase");

            // Try to access derived property via reflection (should not exist, but check for completeness)
            var additionalField1Prop = result.GetType().GetProperty("AdditionalField1");
            if (additionalField1Prop != null)
            {
                var value = additionalField1Prop.GetValue(result);
                value.Should().Be("present-in-derived");
            }
            else
            {
                // This is expected since the result is base type
                // Optionally, you can log or assert inconclusive here
            }
        }
        // Helper derived type for testing
        public class DerivedForEdgeTest : ScoDictionaryNew<string, int>
        {
            public string AdditionalField1 { get; set; }
        }
                
        [TestMethod]
        public void ReadJson_WhenDeserializingToBaseType_TriggersManualCopy()
        {                            
            var mockApplication = new Mock<Microsoft.Office.Interop.Outlook.Application>();
            var globals = new Gsync.AppGlobals(mockApplication.Object) { FS = new AppFileSystemFolderPaths() };

            // Arrange: create and initialize a derived dictionary
            var derived = new ScoDictionaryConverterIntegrationTests.TestDerived().Init(globals);

            // Serialize using your settings (add untyped ScoDictionaryConverter)
            var settings = ScoDictionaryConverterIntegrationTests.TestDerived.GetJsonSettings(globals);            
            settings.Converters.Add(new ScoDictionaryConverter<TestDerived, string, int>());            

            var json = derived.SerializeToString();

            // Act: Deserialize to BASE type, not derived
            var baseResult = JsonConvert.DeserializeObject<ScoDictionaryNew<string, int>>(json, settings);

            // Assert: 
            // 1. The result should NOT be null
            baseResult.Should().NotBeNull();

            // 2. It should NOT be assignable to the derived type (else branch triggers)
            baseResult.Should().NotBeOfType<ScoDictionaryConverterIntegrationTests.TestDerived>();

            // 3. The dictionary content should still be present
            baseResult.ContainsKey("key1").Should().BeTrue();
            baseResult["key1"].Should().Be(1);
            baseResult.ContainsKey("key2").Should().BeTrue();
            baseResult["key2"].Should().Be(2);

            // 4. Optionally: You can check for other property loss if base type does not have them
        }

    }
}
