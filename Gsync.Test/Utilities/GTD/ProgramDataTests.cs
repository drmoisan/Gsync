using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Threading;
using FluentAssertions;
using Gsync.Utilities.GTD;
using Gsync.Utilities.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace Gsync.Test.Utilities.GTD
{
    [TestClass]
    public class ProgramDataTests
    {
        [TestMethod]
        public void AddOrUpdate_ShouldAddAndUpdateValues()
        {
            var data = new ProgramData();
            
            //// Add            
            var elementObject = new ProjectElement();
            elementObject.ID = "1";
            elementObject.Name = "Test";
            data.AddOrUpdate(elementObject, 1, (k, v) => v + 1).Should().Be(1);
            data.TryGetValue(elementObject, out var value).Should().BeTrue();
            value.Should().Be(1);
            //data[elementObject].Should().Be(1);

            // Update
            data.AddOrUpdate(elementObject, 1, (k, v) => v + 1).Should().Be(2);
            data[elementObject].Should().Be(2);
        }

        [TestMethod]
        public void GetOrAdd_ShouldReturnExistingOrAddNew()
        {
            var data = new ProgramData();
            var element = new ProjectElement() { ID = "2", Name = "Test2"};
            
            // Add new
            data.GetOrAdd(element, 5).Should().Be(5);
            data[element].Should().Be(5);

            // Get existing
            data.GetOrAdd(element, 10).Should().Be(5);
        }

        [TestMethod]
        public void TryAddAndTryRemove_ShouldWorkAsExpected()
        {
            var data = new ProgramData();
            var element = new ProjectElement() { ID = "3", Name = "Test3" };
            
            data.TryAdd(element, 7).Should().BeTrue();
            data.TryAdd(element, 8).Should().BeFalse();

            data.TryRemove(element, out var removedValue).Should().BeTrue();
            removedValue.Should().Be(7);
            data.ContainsKey(element).Should().BeFalse();
        }

        [TestMethod]
        public void TryUpdate_ShouldUpdateOnlyIfValueMatches()
        {
            var data = new ProgramData();
            var element = new ProjectElement() { ID = "4", Name = "Test4" };
            
            data[element] = 10;
            data.TryUpdate(element, 20, 10).Should().BeTrue();
            data[element].Should().Be(20);
            data.TryUpdate(element, 30, 10).Should().BeFalse();
        }

        [TestMethod]
        public void Clear_ShouldRemoveAllEntries()
        {
            var data = new ProgramData();
            var element1 = new Mock<IProjectElement>();
            var element2 = new Mock<IProjectElement>();
            element1.SetupGet(e => e.ID).Returns("5");
            element1.SetupGet(e => e.Name).Returns("Test5");
            element2.SetupGet(e => e.ID).Returns("6");
            element2.SetupGet(e => e.Name).Returns("Test6");

            data[element1.Object] = 1;
            data[element2.Object] = 2;
            data.Count.Should().Be(2);

            data.Clear();
            data.Count.Should().Be(0);
        }

        [TestMethod]
        public void Enumeration_ShouldReturnAllEntries()
        {
            var data = new ProgramData();
            // NOTE: Do not use Moq for dictionary keys if mocking an interface. Equals and GetHashCode cannot be reliably overridden.
            
            var object1 = new ProjectElement { ID = "7", Name = "Test7" };
            var object2 = new ProjectElement { ID = "8", Name = "Test8" };

            data[object1] = 11;
            data[object2] = 22;

            //var entries = data.ToList();
            foreach (var entry in data)
            {
                Console.WriteLine($"Key: {entry.Key}, Value: {entry.Value}");
                if (ReferenceEquals(object1, entry.Key))
                {
                    Console.WriteLine($"{entry.Key} matches object1 reference");
                    Console.WriteLine($"Key.Equals(object1): {entry.Key.Equals(object1)}");
                    Console.WriteLine($"Value == 11: {entry.Value == 11}");
                }
                else if (ReferenceEquals(object2, entry.Key))
                {
                    Console.WriteLine($"{entry.Key} matches object2 reference");
                    Console.WriteLine($"Key.Equals(object2): {entry.Key.Equals(object2)}");
                    Console.WriteLine($"Value == 22: {entry.Value == 22}");
                }
                else
                {
                    Console.WriteLine($"{entry.Key} does not match reference for either object");
                }
                Console.WriteLine("");
            }
            Console.WriteLine("");
            data.Should().HaveCount(2);
            data.Should().Contain(x => x.Key.Equals(object1) && x.Value == 11);
            data.Should().Contain(x => x.Key.Equals(object2) && x.Value == 22);
        }

        [TestMethod]
        public void ThreadSafety_ConcurrentAddOrUpdate_ShouldNotThrow()
        {
            var data = new ProgramData();
            var element = new ProjectElement() { ID = "9", Name = "Test9" };

            int exceptions = 0;
            var threads = new Thread[10];
            for (int i = 0; i < threads.Length; i++)
            {
                threads[i] = new Thread(() =>
                {
                    try
                    {
                        for (int j = 0; j < 100; j++)
                        {
                            data.AddOrUpdate(element, 1, (k, v) => v + 1);
                        }
                    }
                    catch
                    {
                        Interlocked.Increment(ref exceptions);
                    }
                });
            }

            foreach (var t in threads) t.Start();
            foreach (var t in threads) t.Join();

            exceptions.Should().Be(0);
            data[element].Should().Be(1 + 10 * 100 - 1);
        }
    }
}