using System.Threading;
using Gsync.Utilities.GTD;
using Gsync.Utilities.Interfaces;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace Gsync.Test.Utilities.GTD
{
    [TestClass]
    public class ProjectElementTests
    {
        [TestMethod]
        public void ID_SetAndGet_ShouldReturnExpectedValue()
        {
            var element = new ProjectElement();
            element.ID = "abc";
            element.ID.Should().Be("abc");
        }

        [TestMethod]
        public void Name_SetAndGet_ShouldReturnExpectedValue()
        {
            var element = new ProjectElement();
            element.Name = "TestName";
            element.Name.Should().Be("TestName");
        }

        [TestMethod]
        public void SettingSameID_ShouldNotRaisePropertyChanged()
        {
            var element = new ProjectElement { ID = "id1" };
            bool raised = false;
            element.PropertyChanged += (s, e) => raised = true;
            element.ID = "id1";
            raised.Should().BeFalse();
        }

        [TestMethod]
        public void SettingSameName_ShouldNotRaisePropertyChanged()
        {
            var element = new ProjectElement { Name = "name1" };
            bool raised = false;
            element.PropertyChanged += (s, e) => raised = true;
            element.Name = "name1";
            raised.Should().BeFalse();
        }

        [TestMethod]
        public void SettingID_ShouldRaisePropertyChanged()
        {
            var element = new ProjectElement();
            string propertyName = null;
            element.PropertyChanged += (s, e) => propertyName = e.PropertyName;
            element.ID = "newId";
            propertyName.Should().Be("ID");
        }

        [TestMethod]
        public void SettingName_ShouldRaisePropertyChanged()
        {
            var element = new ProjectElement();
            string propertyName = null;
            element.PropertyChanged += (s, e) => propertyName = e.PropertyName;
            element.Name = "newName";
            propertyName.Should().Be("Name");
        }

        [TestMethod]
        public void CompareTo_ShouldReturnZeroForEqualNames()
        {
            var a = new ProjectElement { Name = "Alpha" };
            var b = new ProjectElement { Name = "Alpha" };
            a.CompareTo(b).Should().Be(0);
        }

        [TestMethod]
        public void CompareTo_ShouldReturnNegativeForAlphabeticallyBefore()
        {
            var a = new ProjectElement { Name = "Alpha" };
            var b = new ProjectElement { Name = "Beta" };
            a.CompareTo(b).Should().BeLessThan(0);
        }

        [TestMethod]
        public void CompareTo_ShouldReturnPositiveForAlphabeticallyAfter()
        {
            var a = new ProjectElement { Name = "Beta" };
            var b = new ProjectElement { Name = "Alpha" };
            a.CompareTo(b).Should().BeGreaterThan(0);
        }

        [TestMethod]
        public void CompareTo_ShouldReturnOneWhenOtherIsNull()
        {
            var a = new ProjectElement { Name = "Alpha" };
            a.CompareTo(null).Should().Be(1);
        }

        [TestMethod]
        public void Equals_ShouldReturnTrueForSameID()
        {
            var a = new ProjectElement { ID = "1" };
            var b = new ProjectElement { ID = "1" };
            a.Equals(b).Should().BeTrue();
        }

        [TestMethod]
        public void Equals_ShouldReturnFalseForDifferentID()
        {
            var a = new ProjectElement { ID = "1" };
            var b = new ProjectElement { ID = "2" };
            a.Equals(b).Should().BeFalse();
        }

        [TestMethod]
        public void Equals_Object_ShouldReturnTrueForSameID()
        {
            var a = new ProjectElement { ID = "1" };
            object b = new ProjectElement { ID = "1" };
            a.Equals(b).Should().BeTrue();
        }

        [TestMethod]
        public void Equals_Object_ShouldReturnFalseForDifferentID()
        {
            var a = new ProjectElement { ID = "1" };
            object b = new ProjectElement { ID = "2" };
            a.Equals(b).Should().BeFalse();
        }

        [TestMethod]
        public void GetHashCode_ShouldBeConsistentWithID()
        {
            var a = new ProjectElement { ID = "hash" };
            var b = new ProjectElement { ID = "hash" };
            a.GetHashCode().Should().Be(b.GetHashCode());
        }

        [TestMethod]
        public void GetHashCode_ShouldBeZeroWhenIDIsNull()
        {
            var a = new ProjectElement { ID = null };
            a.GetHashCode().Should().Be(0);
        }

        [TestMethod]
        public void ThreadSafety_ConcurrentSetters_ShouldNotThrow()
        {
            var element = new ProjectElement();
            int exceptions = 0;
            var threads = new Thread[10];

            for (int i = 0; i < threads.Length; i++)
            {
                int idx = i;
                threads[i] = new Thread(() =>
                {
                    try
                    {
                        element.ID = idx.ToString();
                        element.Name = "Name" + idx;
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
        }

        [TestMethod]
        public void CompareTo_ShouldWorkWithMockedIProjectElement()
        {
            var element = new ProjectElement { Name = "Alpha" };
            var mock = new Mock<IProjectElement>();
            mock.SetupGet(x => x.Name).Returns("Beta");
            element.CompareTo(mock.Object).Should().BeLessThan(0);
        }

        [TestMethod]
        public void Equals_ShouldWorkWithMockedIProjectElement()
        {
            var element = new ProjectElement { ID = "42" };
            var mock = new Mock<IProjectElement>();
            mock.SetupGet(x => x.ID).Returns("42");
            element.Equals(mock.Object).Should().BeTrue();
        }
    }
}