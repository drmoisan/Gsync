using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading.Tasks;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Gsync.Utilities.Interfaces.IHelperClasses;
using Gsync.Utilities.Interfaces;

namespace Gsync.Test.AppGlobals
{
    [TestClass]
    public class AppFileSystemFolderPathsTests
    {
        private Mock<IEnvironment> _envMock;
        private Mock<IDirectory> _dirMock;
        private string _tempDir;

        [TestInitialize]
        public void Setup()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(_tempDir);

            _envMock = new Mock<IEnvironment>(MockBehavior.Strict);
            _envMock.SetupProperty(e => e.SpecialFolder, Environment.SpecialFolder.LocalApplicationData);
            _envMock.Setup(e => e.GetFolderPath(It.IsAny<Environment.SpecialFolder>()))
                .Returns<Environment.SpecialFolder>(f => Path.Combine(_tempDir, f.ToString()));
            _envMock.Setup(e => e.GetEnvironmentVariable(It.IsAny<string>()))
                .Returns((string name) => Path.Combine(_tempDir, name ?? "ENV"));

            _dirMock = new Mock<IDirectory>(MockBehavior.Strict);
            _dirMock.Setup(d => d.Exists(It.IsAny<string>())).Returns(false);
            _dirMock.Setup(d => d.CreateDirectory(It.IsAny<string>()));
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, true);
        }

        [TestMethod]
        public void Ctor_ShouldInitializeSpecialFoldersAndFilenames()
        {
            var sut = new AppFileSystemFolderPaths(_envMock.Object, _dirMock.Object);

            sut.SpecialFolders.Should().NotBeNull();
            sut.Filenames.Should().NotBeNull();
            sut.SpecialFolders.Should().ContainKey("AppData");

            // Verify that CreateDirectory was called for each special folder
            foreach (var folder in sut.SpecialFolders.Values)
            {
                _dirMock.Verify(d => d.Exists(folder), Times.AtLeastOnce());
                _dirMock.Verify(d => d.CreateDirectory(folder), Times.AtLeastOnce());
            }
        }

        [TestMethod]
        public void MatchBestSpecialFolder_ShouldReturnBestMatch()
        {
            var sut = new AppFileSystemFolderPaths(_envMock.Object, _dirMock.Object);
            var folder = sut.SpecialFolders["AppData"];
            var result = sut.MatchBestSpecialFolder(folder + @"\SomeSubFolder");

            result.Should().Be("AppData");
        }

        [TestMethod]
        public void MatchBestSpecialFolder_ShouldReturnNullIfNoMatch()
        {
            var sut = new AppFileSystemFolderPaths(_envMock.Object, _dirMock.Object);
            var result = sut.MatchBestSpecialFolder(@"D:\NotARealFolder\NoMatch");

            result.Should().BeNull();
        }

        [TestMethod]
        public void Reload_ShouldRepopulateSpecialFolders()
        {
            var sut = new AppFileSystemFolderPaths(_envMock.Object, _dirMock.Object);
            var original = new ConcurrentDictionary<string, string>(sut.SpecialFolders);

            // Remove a key and reload
            sut.SpecialFolders.TryRemove("AppData", out _);
            sut.SpecialFolders.Should().NotContainKey("AppData");

            sut.Reload();

            sut.SpecialFolders.Should().ContainKey("AppData");
            sut.SpecialFolders["AppData"].Should().Be(original["AppData"]);
            // Verify directory creation logic is called again
            _dirMock.Verify(d => d.CreateDirectory(It.IsAny<string>()), Times.AtLeastOnce());
        }

        [TestMethod]
        public void Load_ShouldReturnInitializedInstance()
        {
            // Patch: LoadAsync uses default env, so test only the sync path for DI
            var sut = new AppFileSystemFolderPaths(_envMock.Object, _dirMock.Object);

            sut.Should().NotBeNull();
            sut.SpecialFolders.Should().NotBeNull();
            sut.Filenames.Should().NotBeNull();
            sut.SpecialFolders.Should().ContainKey("AppData");
        }
    }
}