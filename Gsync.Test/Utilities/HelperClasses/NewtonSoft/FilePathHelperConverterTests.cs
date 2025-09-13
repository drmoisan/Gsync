using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using Newtonsoft.Json;
using FluentAssertions;
using Gsync.Utilities.HelperClasses.NewtonSoft;
using Gsync.Utilities.Interfaces;
using Gsync.Utilities.HelperClasses;

namespace Gsync.Test.Utilities.HelperClasses.NewtonSoft
{
    [TestClass]
    public class FilePathHelperConverterTests
    {
        private MockRepository mockRepository;
        private Mock<IFileSystemFolderPaths> mockFileSystemFolderPaths;
        private Mock<JsonReader> mockJsonReader;

        [TestInitialize]
        public void TestInitialize()
        {
            Console.SetOut(new DebugTextWriter());
            this.mockRepository = new MockRepository(MockBehavior.Loose);
            this.mockFileSystemFolderPaths = this.mockRepository.Create<IFileSystemFolderPaths>();
            this.mockJsonReader = this.mockRepository.Create<JsonReader>();
        }

        private FilePathHelperConverter CreateFilePathHelperConverter()
        {
            return new FilePathHelperConverter(
                this.mockFileSystemFolderPaths.Object);
        }

        [TestMethod]
        public void ReadPropertyName_ValidInput_Success()
        {
            // Arrange
            var expected = "property";
            var filePathHelperConverter = this.CreateFilePathHelperConverter();
            //mockJsonReader.Setup(x => x.Read()).Returns(true);
            mockJsonReader.Setup(x => x.TokenType).Returns(JsonToken.PropertyName);
            mockJsonReader.Setup(x => x.Value).Returns(expected);

            // Act
            var actual = filePathHelperConverter.ReadPropertyName(mockJsonReader.Object);

            // Assert
            actual.Should().BeEquivalentTo(expected);
            this.mockJsonReader.Verify(x => x.TokenType, Times.Once());
            this.mockJsonReader.Verify(x => x.Value, Times.Once());
        }

        [ExpectedException(typeof(JsonReaderException))]
        [TestMethod]
        public void ReadPropertyName_NullValue_Failure()
        {
            // Arrange
            string expected = null;
            var filePathHelperConverter = this.CreateFilePathHelperConverter();
            mockJsonReader.Setup(x => x.Read()).Returns(true);
            mockJsonReader.Setup(x => x.TokenType).Returns(JsonToken.PropertyName);
            mockJsonReader.Setup(x => x.Value).Returns(expected);

            // Act
            var actual = filePathHelperConverter.ReadPropertyName(mockJsonReader.Object);

            // Assert
            this.mockJsonReader.Verify(x => x.Read(), Times.Once());
        }

        [ExpectedException(typeof(JsonReaderException))]
        [TestMethod]
        public void ReadPropertyName_WrongType_Failure()
        {
            // Arrange
            string expected = "property";
            var filePathHelperConverter = this.CreateFilePathHelperConverter();
            mockJsonReader.Setup(x => x.Read()).Returns(true);
            mockJsonReader.Setup(x => x.TokenType).Returns(JsonToken.Boolean);
            mockJsonReader.Setup(x => x.Value).Returns(expected);

            // Act
            var actual = filePathHelperConverter.ReadPropertyName(mockJsonReader.Object);

            // Assert
            this.mockJsonReader.Verify(x => x.Read(), Times.Once());
        }

        [TestMethod]
        public void ReadPropertyValue_ValidInput_Success()
        {
            // Arrange
            var expected = "value";
            var filePathHelperConverter = this.CreateFilePathHelperConverter();
            //mockJsonReader.Setup(x => x.Read()).Returns(true);
            mockJsonReader.Setup(x => x.TokenType).Returns(JsonToken.String);
            mockJsonReader.Setup(x => x.Value).Returns(expected);

            // Act
            var actual = filePathHelperConverter.ReadPropertyValue(mockJsonReader.Object);

            // Assert
            actual.Should().BeEquivalentTo(expected);
            this.mockJsonReader.Verify(x => x.TokenType, Times.Once());
            this.mockJsonReader.Verify(x => x.Value, Times.Once());

        }

        [ExpectedException(typeof(JsonReaderException))]
        [TestMethod]
        public void ReadPropertyValue_WrongType_Failure()
        {
            // Arrange
            string expected = "property";
            var filePathHelperConverter = this.CreateFilePathHelperConverter();
            mockJsonReader.Setup(x => x.Read()).Returns(true);
            mockJsonReader.Setup(x => x.TokenType).Returns(JsonToken.Boolean);
            mockJsonReader.Setup(x => x.Value).Returns(expected);

            // Act
            var actual = filePathHelperConverter.ReadPropertyValue(mockJsonReader.Object);

            // Assert
            this.mockJsonReader.Verify(x => x.Read(), Times.Once());
        }

        //[TestMethod]
        //public void ReadJson_StateUnderTest_ExpectedBehavior()
        //{
        //    // Arrange
        //    var filePathHelperConverter = this.CreateFilePathHelperConverter();
        //    JsonReader reader = null;
        //    Type objectType = null;
        //    FilePathHelper existingValue = null;
        //    bool hasExistingValue = false;
        //    JsonSerializer serializer = null;

        //    // Act
        //    var result = filePathHelperConverter.ReadJson(
        //        reader,
        //        objectType,
        //        existingValue,
        //        hasExistingValue,
        //        serializer);

        //    // Assert
        //    Assert.Fail();
        //    this.mockRepository.VerifyAll();
        //}

        //[TestMethod]
        //public void WriteJson_StateUnderTest_ExpectedBehavior()
        //{
        //    // Arrange
        //    var filePathHelperConverter = this.CreateFilePathHelperConverter();
        //    JsonWriter writer = null;
        //    FilePathHelper value = null;
        //    JsonSerializer serializer = null;

        //    // Act
        //    filePathHelperConverter.WriteJson(
        //        writer,
        //        value,
        //        serializer);

        //    // Assert
        //    Assert.Fail();
        //    this.mockRepository.VerifyAll();
        //}
    }
}
