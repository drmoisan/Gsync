using Gsync.Utilities.Interfaces;
using Gsync.Utilities.ReusableTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace Gsync.Test.Utilities.ReusableTypes
{
    [TestClass]
    public class ConcurrentObservableDictionaryTests
    {
        private ConcurrentObservableDictionary<string, int> _dictionary;

        [TestInitialize]
        public void Setup()
        {
            _dictionary = new ConcurrentObservableDictionary<string, int>();
        }

        [TestMethod]
        public void AddOrUpdate_AddsNewItem()
        {
            // Arrange
            string key = "key1";
            int value = 1;

            // Act
            _dictionary.AddOrUpdate(key, value);

            // Assert
            Assert.AreEqual(value, _dictionary[key]);
        }

        [TestMethod]
        public void AddOrUpdate_UpdatesExistingItem()
        {
            // Arrange
            string key = "key1";
            int initialValue = 1;
            int updatedValue = 2;
            _dictionary.TryAdd(key, initialValue);

            // Act
            _dictionary.AddOrUpdate(key, updatedValue);

            // Assert
            Assert.AreEqual(updatedValue, _dictionary[key]);
        }

        [TestMethod]
        public void GetOrAdd_AddsNewItem()
        {
            // Arrange
            string key = "key1";
            int value = 1;

            // Act
            var result = _dictionary.GetOrAdd(key, value);

            // Assert
            Assert.AreEqual(value, result);
            Assert.AreEqual(value, _dictionary[key]);
        }

        [TestMethod]
        public void GetOrAdd_ReturnsExistingItem()
        {
            // Arrange
            string key = "key1";
            int initialValue = 1;
            _dictionary.TryAdd(key, initialValue);

            // Act
            var result = _dictionary.GetOrAdd(key, 2);

            // Assert
            Assert.AreEqual(initialValue, result);
        }

        [TestMethod]
        public void TryAdd_AddsNewItem()
        {
            // Arrange
            string key = "key1";
            int value = 1;

            // Act
            var result = _dictionary.TryAdd(key, value);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(value, _dictionary[key]);
        }

        [TestMethod]
        public void TryAdd_DoesNotAddExistingItem()
        {
            // Arrange
            string key = "key1";
            int initialValue = 1;
            _dictionary.TryAdd(key, initialValue);

            // Act
            var result = _dictionary.TryAdd(key, 2);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual(initialValue, _dictionary[key]);
        }

        [TestMethod]
        public void TryRemove_RemovesExistingItem()
        {
            // Arrange
            string key = "key1";
            int value = 1;
            _dictionary.TryAdd(key, value);

            // Act
            var result = _dictionary.TryRemove(key, out var removedValue);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(value, removedValue);
            Assert.IsFalse(_dictionary.ContainsKey(key));
        }

        [TestMethod]
        public void TryRemove_DoesNotRemoveNonExistingItem()
        {
            // Arrange
            string key = "key1";

            // Act
            var result = _dictionary.TryRemove(key, out var removedValue);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual(default(int), removedValue);
        }

        [TestMethod]
        public void TryUpdate_UpdatesExistingItem()
        {
            // Arrange
            string key = "key1";
            int initialValue = 1;
            int updatedValue = 2;
            _dictionary.TryAdd(key, initialValue);

            // Act
            var result = _dictionary.TryUpdate(key, updatedValue, initialValue);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual(updatedValue, _dictionary[key]);
        }

        [TestMethod]
        public void TryUpdate_DoesNotUpdateIfComparisonFails()
        {
            // Arrange
            string key = "key1";
            int initialValue = 1;
            int updatedValue = 2;
            _dictionary.TryAdd(key, initialValue);

            // Act
            var result = _dictionary.TryUpdate(key, updatedValue, 3);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual(initialValue, _dictionary[key]);
        }

        [TestMethod]
        public void Clear_RemovesAllItems()
        {
            // Arrange
            _dictionary.TryAdd("key1", 1);
            _dictionary.TryAdd("key2", 2);

            // Act
            _dictionary.Clear();

            // Assert
            Assert.AreEqual(0, _dictionary.Count);
        }

        [TestMethod]
        public void AddPartialObserver_AddsObserverForKey()
        {
            // Arrange
            string key = "key1";
            var observer = new TestObserver<string, int>();

            // Act
            _dictionary.AddPartialObserver(observer, key);
            _dictionary.AddOrUpdate(key, 1);

            // Assert
            Assert.AreEqual(1, observer.ReceivedEvents.Count);
        }

        [TestMethod]
        public void RemovePartialObserver_RemovesObserverForKey()
        {
            // Arrange
            string key = "key1";
            var observer = new TestObserver<string, int>();
            _dictionary.AddPartialObserver(observer, key);

            // Act
            _dictionary.RemovePartialObserver(observer, key);
            _dictionary.AddOrUpdate(key, 1);

            // Assert
            Assert.AreEqual(0, observer.ReceivedEvents.Count);
        }

        [TestMethod]
        public void RemoveAllObservers_RemovesAllObservers()
        {
            // Arrange
            string key = "key1";
            var observer = new TestObserver<string, int>();
            _dictionary.AddPartialObserver(observer, key);

            // Act
            _dictionary.RemoveAllObservers();
            _dictionary.AddOrUpdate(key, 1);

            // Assert
            Assert.AreEqual(0, observer.ReceivedEvents.Count);
        }

        private class TestObserver<TKey, TValue> : IDictionaryObserver<TKey, TValue>
        {
            public List<DictionaryChangedEventArgs<TKey, TValue>> ReceivedEvents { get; } = new List<DictionaryChangedEventArgs<TKey, TValue>>();

            public void OnEventOccur(DictionaryChangedEventArgs<TKey, TValue> args)
            {
                ReceivedEvents.Add(args);
            }
        }
    }
}
