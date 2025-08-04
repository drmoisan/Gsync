using FluentAssertions;
using Gsync.OutlookInterop.Item;
using Gsync.OutlookInterop.Interfaces.Items;
using Gsync.Utilities.HelperClasses;
using log4net.Appender;
using log4net.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Gsync.Test.OutlookInterop.Item
{
    [TestClass]
    public class OutlookItemWrapperTests
    {
        private Mock<IItem> _itemMock;
        private IItem _item;
        private Mock<ItemEvents_10_Event> _eventsMock;
        private ItemEvents_10_Event _events;
        private OutlookItemWrapper _wrapper;
        private DynamicIItem _dynItem;
        private ImmutableHashSet<string> _supportedTypes;

        [TestInitialize]
        public void Setup()
        {
            Console.SetOut(new DebugTextWriter());
            
            _itemMock = CreateIItemMock();            
            _eventsMock = new Mock<ItemEvents_10_Event>();
            _dynItem = new DynamicIItem();
        }

        private Mock<IItem> CreateIItemMock()
        {
            var mock = new Mock<IItem>();
            mock.SetupAllProperties();
            return mock;
        }

        private OutlookItemWrapper CreateWrapper(IItem item, ItemEvents_10_Event events, ImmutableHashSet<string> supportedTypes)
        {
            return new TestableOutlookItemWrapper(item, events, supportedTypes);
        }

        private OutlookItemWrapper CreateWrapper(DynamicIItem dyn, IItem item, ItemEvents_10_Event events, ImmutableHashSet<string> supportedTypes)
        {
            var wrapper = (OutlookItemWrapper)Activator.CreateInstance(
                typeof(OutlookItemWrapper),
                BindingFlags.NonPublic | BindingFlags.Instance,
                null,
                new object[] { item, events, supportedTypes },
                null
            );
            var dynField = typeof(OutlookItemWrapper).GetField("_dyn", BindingFlags.NonPublic | BindingFlags.Instance);
            dynField.SetValue(wrapper, dyn);
            return wrapper;
        }

        private OutlookItemWrapper CreateWrapper(Mock<IItem> mock)
        {            
            return new OutlookItemWrapper(mock.Object);
        }

        private bool ProtectedIsComObjectFunc(object obj) 
        {
            var wrapper = CreateWrapper(_dynItem, _itemMock.Object, _eventsMock.Object, _supportedTypes);
            var protectedIsComObjectFunc = wrapper.GetType()
                .GetMethod(
                "IsComObjectFunc",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic);

            if (protectedIsComObjectFunc is null) { throw new InvalidOperationException("Method IsComObjectFunc not found"); }

            var result = protectedIsComObjectFunc.Invoke(wrapper, new object[] { obj });
            if (result is null) { throw new InvalidOperationException("Method IsComObjectFunc returned null"); }
            return (bool)result;
        }

        private void LockMockObjects()
        {
            _item = _itemMock.Object;
            _events = _eventsMock.Object;
            _supportedTypes = ImmutableHashSet.Create(_item.GetType().Name);
            _wrapper = CreateWrapper(_item, _events, _supportedTypes);
        }


        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void _OutlookItemWrapper_ExceptionIfNullObject()
        {
            Console.WriteLine("Testing OutlookItemWrapper with null object");
            var wrapper = new OutlookItemWrapper(null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void _OutlookItemWrapper_ExceptionIfUnsupportedType()
        {
            Console.WriteLine("Testing OutlookItemWrapper with unsupported type");
            var wrapper = new OutlookItemWrapper(new object());                        
        }

        [TestMethod]
        public void Application_Property_ReturnsValue()
        {
            // Arrange
            var app = new Mock<Application>().Object;
            _itemMock.Setup(m => m.Application).Returns(app);
            LockMockObjects();
            
            // Act
            var app2 = _wrapper.Application;

            // Assert
            Assert.AreEqual(app, _wrapper.Application);
        }

        [TestMethod]
        public void Attachments_Property_ReturnsValue()
        {            
            var attachments = new Mock<Attachments>().Object;
            _itemMock.Setup(m => m.Attachments).Returns(attachments);
            LockMockObjects();
            Assert.AreEqual(attachments, _wrapper.Attachments);
        }

        [TestMethod]
        public void BillingInformation_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.BillingInformation, "Test");
            LockMockObjects();            
            Assert.AreEqual("Test", _wrapper.BillingInformation);
            _wrapper.BillingInformation = "NewValue";
            Assert.AreEqual("NewValue", _wrapper.BillingInformation);
        }

        [TestMethod]
        public void Body_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.Body, "TestBody");
            LockMockObjects();
            Assert.AreEqual("TestBody", _wrapper.Body);
            _wrapper.Body = "NewBody";
            Assert.AreEqual("NewBody", _wrapper.Body);
        }

        [TestMethod]
        public void Categories_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.Categories, "TestCat");
            LockMockObjects();
            Assert.AreEqual("TestCat", _wrapper.Categories);
            _wrapper.Categories = "NewCat";
            Assert.AreEqual("NewCat", _wrapper.Categories);
        }

        [TestMethod]
        public void Class_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.Class).Returns(OlObjectClass.olMail);
            LockMockObjects();
            Assert.AreEqual(OlObjectClass.olMail, _wrapper.Class);
        }

        [TestMethod]
        public void Companies_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.Companies, "TestCo");
            LockMockObjects();
            Assert.AreEqual("TestCo", _wrapper.Companies);
            _wrapper.Companies = "NewCo";
            Assert.AreEqual("NewCo", _wrapper.Companies);
        }

        [TestMethod]
        public void ConversationID_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.ConversationID).Returns("ConvId");
            LockMockObjects();
            Assert.AreEqual("ConvId", _wrapper.ConversationID);
        }

        [TestMethod]
        public void CreationTime_Property_ReturnsValue()
        {
            var dt = DateTime.Now;
            _itemMock.Setup(m => m.CreationTime).Returns(dt);
            LockMockObjects();
            Assert.AreEqual(dt, _wrapper.CreationTime);
        }

        [TestMethod]
        public void Dispose_UnsubscribesEventsAndReleasesComObject()
        {
            // Update the event handler type to match the expected type            
            var mailItemMock = new Mock<MailItem>();
            var eventsMock = new Mock<ItemEvents_10_Event>();

            bool unsubscribed = false;
            eventsMock.SetupRemove(e => e.AttachmentAdd -= It.IsAny<ItemEvents_10_AttachmentAddEventHandler>())
                .Callback(() => unsubscribed = true);

            var wrapper = new TestableOutlookItemWrapper(mailItemMock.Object, eventsMock.Object);
            wrapper.Dispose();

            Assert.IsTrue(wrapper.ReleasedObjects.Count >= 1, "Should release at least one object.");
            Assert.IsTrue(wrapper.ReleasedObjects.Contains(mailItemMock.Object),
                "MailItem mock should be among the released objects.");
            Assert.IsTrue(unsubscribed, "Events should be unsubscribed");
        }

        [TestMethod]
        public void Dispose_MultipleCalls_IsIdempotent()
        {
            var mailItemMock = new Mock<MailItem>();
            var wrapper = new TestableOutlookItemWrapper(mailItemMock.Object);
            wrapper.Dispose();
            int releasedAfterFirstCall = wrapper.ReleasedObjects.Count;
            wrapper.Dispose(); // second call should have no effect
            int releasedAfterSecondCall = wrapper.ReleasedObjects.Count;

            Assert.AreEqual(releasedAfterFirstCall, releasedAfterSecondCall, "Dispose should not release objects more than once.");
        }

        [TestMethod]
        public void Dispose_NoEvents_NoException()
        {
            var mailItemMock = new Mock<MailItem>();
            var wrapper = new TestableOutlookItemWrapper(mailItemMock.Object);
            wrapper.Dispose();
            Assert.IsTrue(wrapper.ReleasedObjects.Contains(mailItemMock.Object));
        }

        [TestMethod]
        public void EntryID_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.EntryID).Returns("EntryId");
            LockMockObjects();
            Assert.AreEqual("EntryId", _wrapper.EntryID);
        }

        [TestMethod]
        public void Importance_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.Importance, OlImportance.olImportanceNormal);
            LockMockObjects();
            Assert.AreEqual(OlImportance.olImportanceNormal, _wrapper.Importance);
            _wrapper.Importance = OlImportance.olImportanceHigh;
            Assert.AreEqual(OlImportance.olImportanceHigh, _wrapper.Importance);
        }

        //[TestMethod]
        //public void ItemProperties_Property_ReturnsValue()
        //{
        //    var mock = CreateIItemMock();
        //    var itemProps = new Mock<ItemProperties>().Object;
        //    mock.Setup(m => m.ItemProperties).Returns(itemProps);
        //    var _wrapper.= CreateWrapper(mock);
        //    Assert.AreEqual(itemProps, _wrapper.ItemProperties);
        //}

        [TestMethod]
        public void LastModificationTime_Property_ReturnsValue()
        {            
            var dt = DateTime.Now;
            _itemMock.Setup(m => m.LastModificationTime).Returns(dt);
            LockMockObjects();
            Assert.AreEqual(dt, _wrapper.LastModificationTime);
        }

        [TestMethod]
        public void MessageClass_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.MessageClass).Returns("IPM.Note");
            LockMockObjects();
            Assert.AreEqual("IPM.Note", _wrapper.MessageClass);
        }

        [TestMethod]
        public void Mileage_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.Mileage, "TestMileage");
            LockMockObjects();
            Assert.AreEqual("TestMileage", _wrapper.Mileage);
            _wrapper.Mileage = "NewMileage";
            Assert.AreEqual("NewMileage", _wrapper.Mileage);
        }

        [TestMethod]
        public void NoAging_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.NoAging, false);
            LockMockObjects();
            Assert.IsFalse(_wrapper.NoAging);
            _wrapper.NoAging = true;
            Assert.IsTrue(_wrapper.NoAging);
        }

        [TestMethod]
        public void OutlookInternalVersion_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.OutlookInternalVersion).Returns(1234);
            LockMockObjects();
            Assert.AreEqual(1234, _wrapper.OutlookInternalVersion);
        }

        [TestMethod]
        public void OutlookVersion_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.OutlookVersion).Returns("16.0");
            LockMockObjects();
            Assert.AreEqual("16.0", _wrapper.OutlookVersion);
        }

        [TestMethod]
        public void Parent_Property_ReturnsValue()
        {            
            var parent = new object();
            _itemMock.Setup(m => m.Parent).Returns(parent);
            LockMockObjects();
            Assert.AreEqual(parent, _wrapper.Parent);
        }

        [TestMethod]
        public void Saved_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.Saved).Returns(true);
            LockMockObjects();
            Assert.IsTrue(_wrapper.Saved);
        }

        [TestMethod]
        public void Sensitivity_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.Sensitivity, OlSensitivity.olNormal);
            LockMockObjects();
            Assert.AreEqual(OlSensitivity.olNormal, _wrapper.Sensitivity);
            _wrapper.Sensitivity = OlSensitivity.olConfidential;
            Assert.AreEqual(OlSensitivity.olConfidential, _wrapper.Sensitivity);
        }

        [TestMethod]
        public void Session_Property_ReturnsValue()
        {            
            var session = new Mock<NameSpace>().Object;
            _itemMock.Setup(m => m.Session).Returns(session);
            LockMockObjects();
            Assert.AreEqual(session, _wrapper.Session);
        }

        [TestMethod]
        public void Size_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.Size).Returns(42);
            LockMockObjects();
            Assert.AreEqual(42, _wrapper.Size);
        }

        [TestMethod]
        public void Subject_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.Subject, "TestSubject");
            LockMockObjects();
            Assert.AreEqual("TestSubject", _wrapper.Subject);
            _wrapper.Subject = "NewSubject";
            Assert.AreEqual("NewSubject", _wrapper.Subject);
        }

        [TestMethod]
        public void UnRead_Property_GetSet()
        {
            _itemMock.SetupProperty(m => m.UnRead, false);
            LockMockObjects();
            Assert.IsFalse(_wrapper.UnRead);
            _wrapper.UnRead = true;
            Assert.IsTrue(_wrapper.UnRead);
        }

        [TestMethod]
        public void Close_Method_CallsUnderlying()
        {
            _itemMock.Setup(m => m.Close(It.IsAny<OlInspectorClose>())).Verifiable();
            LockMockObjects();
            _wrapper.Close(OlInspectorClose.olSave);
            _itemMock.Verify(m => m.Close(OlInspectorClose.olSave), Times.Once);
        }

        [TestMethod]
        public void Close_WhenItemNull_Throws()
        {
            // Arrange
            var dummy = new DummyWithWrongClose();

            var dummyTypeName = dummy.GetType().Name;
            var supportedTypes = new HashSet<string>
            {
                dummyTypeName
            }.ToImmutableHashSet();
            var mockEvents = new Mock<ItemEvents_10_Event>();

            var wrapper = new TestableOutlookItemWrapper(dummy, mockEvents.Object, supportedTypes);
            wrapper.Set_item(null); // Set item to null to simulate the condition

            // Act
            System.Action act = () => wrapper.Close(OlInspectorClose.olDiscard);

            // Assert: 
            act.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void Close_WhenMethodNotFound_Throws()
        {
            // Arrange
            var dummy = new DummyWithWrongClose();

            var dummyTypeName = dummy.GetType().Name;
            var supportedTypes = new HashSet<string>
            {
                dummyTypeName
            }.ToImmutableHashSet();
            var mockEvents = new Mock<ItemEvents_10_Event>();

            var wrapper = new TestableOutlookItemWrapper(dummy, mockEvents.Object, supportedTypes);

            // Act
            System.Action act = () => wrapper.Close(OlInspectorClose.olDiscard);

            // Assert: 
            act.Should().Throw<InvalidOperationException>();
            //act.Should().Throw<TargetInvocationException>().Which
            //    .InnerException.Should().BeOfType<InvalidOperationException>();
        }

        [TestMethod]
        public void Copy_Method_CallsUnderlying()
        {            
            var copyObj = new object();
            _itemMock.Setup(m => m.Copy()).Returns(copyObj);
            LockMockObjects();            
            Assert.AreEqual(copyObj, _wrapper.Copy());
        }

        [TestMethod]
        public void Delete_Method_CallsUnderlying()
        {
            _itemMock.Setup(m => m.Delete()).Verifiable();
            LockMockObjects();
            _wrapper.Delete();
            _itemMock.Verify(m => m.Delete(), Times.Once);
        }

        [TestMethod]
        public void Display_Method_CallsUnderlying()
        {
            _itemMock.Setup(m => m.Display(It.IsAny<object>())).Verifiable();
            LockMockObjects();
            _wrapper.Display(true);
            _itemMock.Verify(m => m.Display(true), Times.Once);
        }

        [TestMethod]
        public void Display_CallsUnderlying_WithNullModal()
        {
            // Arrange
            _itemMock.Setup(m => m.Display(null)).Verifiable();
            LockMockObjects();

            // Act
            _wrapper.Display(null);

            // Assert
            _itemMock.Verify(m => m.Display(null), Times.Once);
        }

        [TestMethod]
        public void Move_Method_CallsUnderlying()
        {            
            var folder = new Mock<MAPIFolder>().Object;
            var movedObj = new object();
            _itemMock.Setup(m => m.Move(folder)).Returns(movedObj);
            LockMockObjects();
            Assert.AreEqual(movedObj, _wrapper.Move(folder));
        }

        [TestMethod]
        public void PrintOut_Method_CallsUnderlying()
        {
            _itemMock.Setup(m => m.PrintOut()).Verifiable();
            LockMockObjects();
            _wrapper.PrintOut();
            _itemMock.Verify(m => m.PrintOut(), Times.Once);
        }

        [TestMethod]
        public void Save_Method_CallsUnderlying()
        {
            _itemMock.Setup(m => m.Save()).Verifiable();
            LockMockObjects();
            _wrapper.Save();
            _itemMock.Verify(m => m.Save(), Times.Once);
        }

        [TestMethod]
        public void SaveAs_Method_CallsUnderlying()
        {
            _itemMock.Setup(m => m.SaveAs(It.IsAny<string>(), It.IsAny<object>())).Verifiable();
            LockMockObjects();
            _wrapper.SaveAs("path", "type");
            _itemMock.Verify(m => m.SaveAs("path", "type"), Times.Once);
        }

        [TestMethod]
        public void SaveAs_CallsUnderlying_WithNullType()
        {
            // Arrange
            _itemMock.Setup(m => m.SaveAs(It.IsAny<string>(), null)).Verifiable();
            LockMockObjects();

            // Act
            _wrapper.SaveAs("path", null);

            // Assert
            _itemMock.Verify(m => m.SaveAs("path", null), Times.Once);
        }

        [TestMethod]
        [ExpectedException(typeof(COMException))]
        public void Property_AccessThrowsException_IsPropagated()
        {
            _itemMock.Setup(m => m.Application).Throws(new COMException("fail!"));
            LockMockObjects();            
            var unused = _wrapper.Application;
        }

        [TestMethod]
        [ExpectedException(typeof(COMException))]
        public void Property_SetThrowsException_IsPropagated()
        {
            _itemMock.SetupSet(m => m.Body = It.IsAny<string>()).Throws(new COMException("fail!"));
            LockMockObjects();
            _wrapper.Body = "Should throw";
        }

        [TestMethod]
        public void IsComObjectFunc_ReturnsFalse_ForNull()
        {
            var result = ProtectedIsComObjectFunc(null);            
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsComObjectFunc_ReturnsFalse_ForManagedObject()
        {
            var managedObj = new object();
            var result = ProtectedIsComObjectFunc(managedObj);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsComObjectFunc_ReturnsTrue_ForComObject()
        {            
            // Use a real COM object (Scripting.Dictionary is available on most Windows systems)
            Type comType = Type.GetTypeFromProgID("Scripting.Dictionary");
            if (comType == null)
            {
                Assert.Inconclusive("Scripting.Dictionary COM type not available on this system.");
            }
            else
            {
                object comObj = Activator.CreateInstance(comType);
                try
                {
                    var result = ProtectedIsComObjectFunc(comObj);
                    Assert.IsTrue(result);
                }
                finally
                {
                    Marshal.ReleaseComObject(comObj);
                }
            }
        }

        [TestMethod]
        public void AttachComEvents_DoesNothing_WhenComEventsIsNull()
        {
            // Arrange
            var wrapperType = typeof(OutlookItemWrapper);
            var wrapper = (OutlookItemWrapper)Activator.CreateInstance(
                wrapperType,
                BindingFlags.NonPublic | BindingFlags.Instance,
                null,
                new object[] { _itemMock.Object, null, ImmutableHashSet.Create(_itemMock.Object.GetType().Name) },
                null
            );
            var comEventsField = wrapperType.GetField("_comEvents", BindingFlags.NonPublic | BindingFlags.Instance);
            comEventsField.SetValue(wrapper, null);

            // Act & Assert: Should not throw
            var method = wrapperType.GetMethod("AttachComEvents", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(wrapper, null);
        }

        [TestMethod]
        public void DetachComEvents_DoesNothing_WhenComEventsIsNull()
        {
            // Arrange
            var wrapperType = typeof(OutlookItemWrapper);
            var wrapper = (OutlookItemWrapper)Activator.CreateInstance(
                wrapperType,
                BindingFlags.NonPublic | BindingFlags.Instance,
                null,
                new object[] { _itemMock.Object, null, ImmutableHashSet.Create(_itemMock.Object.GetType().Name) },
                null
            );
            var comEventsField = wrapperType.GetField("_comEvents", BindingFlags.NonPublic | BindingFlags.Instance);
            comEventsField.SetValue(wrapper, null);

            // Act & Assert: Should not throw
            var method = wrapperType.GetMethod("DetachComEvents", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(wrapper, null);
        }

        [TestMethod]
        public void OnAttachmentAdd_RaisesEvent()
        {
            LockMockObjects();
            bool called = false;
            var attachment = new Mock<Attachment>().Object;
            _wrapper.AttachmentAdd += a =>
            {
                called = true;
                Assert.AreEqual(attachment, a);
            };

            (_wrapper as TestableOutlookItemWrapper).InvokeOnAttachmentAdd(attachment);
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnAttachmentRead_RaisesEvent()
        {
            LockMockObjects();
            bool called = false;
            var attachment = new Mock<Attachment>().Object;
            _wrapper.AttachmentRead += a =>
            {
                called = true;
                Assert.AreEqual(attachment, a);
            };

            (_wrapper as TestableOutlookItemWrapper).InvokeOnAttachmentRead(attachment);
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnAttachmentRemove_RaisesEvent()
        {
            LockMockObjects();
            bool called = false;
            var attachment = new Mock<Attachment>().Object;
            _wrapper.AttachmentRemove += a =>
            {
                called = true;
                Assert.AreEqual(attachment, a);
            };

            (_wrapper as TestableOutlookItemWrapper).InvokeOnAttachmentRemove(attachment);
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnBeforeDelete_RaisesEvent()
        {
            LockMockObjects();
            bool called = false;
            object item = new object();
            _wrapper.BeforeDelete += (object i, ref bool c) =>
            {
                called = true;
                Assert.AreEqual(item, i);
                c = true;
            };

            bool cancelArg = false;
            (_wrapper as TestableOutlookItemWrapper).InvokeOnBeforeDelete(item, ref cancelArg);
            Assert.IsTrue(called);
            Assert.IsTrue(cancelArg);
        }

        [TestMethod]
        public void OnCloseEvent_RaisesEvent()
        {
            LockMockObjects();
            bool called = false;
            _wrapper.CloseEvent += (ref bool c) =>
            {
                called = true;
                c = true;
            };

            bool cancel = false;
            (_wrapper as TestableOutlookItemWrapper).InvokeOnCloseEvent(ref cancel);
            Assert.IsTrue(called);
            Assert.IsTrue(cancel);
        }

        [TestMethod]
        public void OnOpen_RaisesEvent()
        {
            LockMockObjects();
            bool called = false;
            _wrapper.Open += (ref bool c) =>
            {
                called = true;
                c = true;
            };

            bool cancel = false;
            (_wrapper as TestableOutlookItemWrapper).InvokeOnOpen(ref cancel);
            Assert.IsTrue(called);
            Assert.IsTrue(cancel);
        }

        [TestMethod]
        public void OnPropertyChange_RaisesEvent()
        {
            LockMockObjects();
            bool called = false;
            string propertyName = "TestProp";
            _wrapper.PropertyChange += name =>
            {
                called = true;
                Assert.AreEqual(propertyName, name);
            };

            (_wrapper as TestableOutlookItemWrapper).InvokeOnPropertyChange(propertyName);
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnRead_RaisesEvent()
        {
            LockMockObjects();
            bool called = false;
            _wrapper.Read += () => called = true;

            (_wrapper as TestableOutlookItemWrapper).InvokeOnRead();
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnWrite_RaisesEvent()
        {
            LockMockObjects();
            bool called = false;
            _wrapper.Write += (ref bool c) =>
            {
                called = true;
                c = true;
            };

            bool cancel = false;
            (_wrapper as TestableOutlookItemWrapper).InvokeOnWrite(ref cancel);
            Assert.IsTrue(called);
            Assert.IsTrue(cancel);
        }

        [TestMethod]
        public void EqualityComparer_Default_IsIItemEqualityComparer()
        {            
            LockMockObjects();
            Assert.IsInstanceOfType(_wrapper.EqualityComparer, typeof(IItemEqualityComparer));
        }

        [TestMethod]
        public void EqualityComparer_CanBeInjected_AndUsed()
        {
            LockMockObjects();
            var customComparer = new Mock<IEqualityComparer<IItem>>();
            customComparer.Setup(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>())).Returns(true);
            customComparer.Setup(c => c.GetHashCode(It.IsAny<IItem>())).Returns(123);

            _wrapper.EqualityComparer = customComparer.Object;
            Assert.AreEqual(customComparer.Object, _wrapper.EqualityComparer);

            // Should use the injected comparer
            Assert.IsTrue(_wrapper.Equals((IItem)null));
            Assert.AreEqual(123, _wrapper.GetHashCode());
        }

        [TestMethod]
        public void EqualsIItem_UsesEqualityComparer()
        {
            LockMockObjects();
            var otherMock = CreateIItemMock();
            var customComparer = new Mock<IEqualityComparer<IItem>>();
            customComparer.Setup(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>())).Returns(false);
            _wrapper.EqualityComparer = customComparer.Object;            
            var otherTypes = ImmutableHashSet.Create(otherMock.Object.GetType().Name);
            var otherWrapper = CreateWrapper(otherMock.Object, _events, otherTypes);

            Assert.IsFalse(_wrapper.Equals(otherWrapper));
            customComparer.Verify(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>()), Times.Once);
        }

        [TestMethod]
        public void EqualsObject_UsesEqualityComparer()
        {
            LockMockObjects();
            var otherMock = CreateIItemMock();
            var customComparer = new Mock<IEqualityComparer<IItem>>();
            customComparer.Setup(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>())).Returns(true);
            _wrapper.EqualityComparer = customComparer.Object;
            var otherTypes = ImmutableHashSet.Create(otherMock.Object.GetType().Name);
            var otherWrapper = CreateWrapper(otherMock.Object, _events, otherTypes);

            Assert.IsTrue(_wrapper.Equals((object)otherWrapper));
            customComparer.Verify(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>()), Times.Once);
        }

        [TestMethod]
        public void GetHashCode_UsesEqualityComparer()
        {            
            var customComparer = new Mock<IEqualityComparer<IItem>>();
            customComparer.Setup(c => c.GetHashCode(It.IsAny<IItem>())).Returns(42);
            LockMockObjects();
            _wrapper.EqualityComparer = customComparer.Object;            

            Assert.AreEqual(42, _wrapper.GetHashCode());
            customComparer.Verify(c => c.GetHashCode(_wrapper), Times.Once);
        }

        [TestMethod]
        public void Actions_Property_ReturnsValue()
        {            
            var actions = new Mock<Actions>().Object;
            _itemMock.Setup(m => m.Actions).Returns(actions);
            LockMockObjects();            
            Assert.AreEqual(actions, _wrapper.Actions);
        }

        [TestMethod]
        public void ConversationIndex_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.ConversationIndex).Returns("conv-index");
            LockMockObjects();
            Assert.AreEqual("conv-index", _wrapper.ConversationIndex);
        }

        [TestMethod]
        public void ConversationTopic_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.ConversationTopic).Returns("conv-topic");
            LockMockObjects();
            Assert.AreEqual("conv-topic", _wrapper.ConversationTopic);
        }

        [TestMethod]
        public void FormDescription_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.FormDescription).Returns("form-desc");
            LockMockObjects();
            Assert.AreEqual("form-desc", _wrapper.FormDescription);
        }

        [TestMethod]
        public void GetInspector_Property_ReturnsValue()
        {            
            var inspector = new object();
            _itemMock.Setup(m => m.GetInspector).Returns(inspector);
            LockMockObjects();
            Assert.AreEqual(inspector, _wrapper.GetInspector);
        }

        [TestMethod]
        public void MAPIOBJECT_Property_ReturnsValue()
        {            
            var mapiObj = new object();
            _itemMock.Setup(m => m.MAPIOBJECT).Returns(mapiObj);
            LockMockObjects();
            Assert.AreEqual(mapiObj, _wrapper.MAPIOBJECT);
        }

        [TestMethod]
        public void UserProperties_Property_ReturnsValue()
        {            
            var userProps = new object();
            _itemMock.Setup(m => m.UserProperties).Returns(userProps);
            LockMockObjects();
            Assert.AreEqual(userProps, _wrapper.UserProperties);
        }

        [TestMethod]
        public void AutoResolvedWinner_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.AutoResolvedWinner).Returns(true);
            LockMockObjects();
            Assert.IsTrue(_wrapper.AutoResolvedWinner);
        }

        [TestMethod]
        public void Conflicts_Property_ReturnsValue()
        {            
            var conflicts = new Mock<Conflicts>().Object;
            _itemMock.Setup(m => m.Conflicts).Returns(conflicts);
            LockMockObjects();
            Assert.AreEqual(conflicts, _wrapper.Conflicts);
        }

        [TestMethod]
        public void DownloadState_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.DownloadState).Returns(OlDownloadState.olFullItem);
            LockMockObjects();
            Assert.AreEqual(OlDownloadState.olFullItem, _wrapper.DownloadState);
        }

        [TestMethod]
        public void IsConflict_Property_ReturnsValue()
        {
            _itemMock.Setup(m => m.IsConflict).Returns(true);
            LockMockObjects();
            Assert.IsTrue(_wrapper.IsConflict);
        }

        [TestMethod]
        public void Links_Property_ReturnsValue()
        {            
            var links = new Mock<Links>().Object;
            _itemMock.Setup(m => m.Links).Returns(links);
            LockMockObjects();
            Assert.AreEqual(links, _wrapper.Links);
        }

        [TestMethod]
        public void PropertyAccessor_Property_ReturnsValue()
        {            
            var accessor = new Mock<PropertyAccessor>().Object;
            _itemMock.Setup(m => m.PropertyAccessor).Returns(accessor);
            LockMockObjects();
            Assert.AreEqual(accessor, _wrapper.PropertyAccessor);
        }

        [TestMethod]
        public void GetRawHeaders_ReturnsHeaders_WhenPropertyAccessorReturnsString()
        {
            // Arrange
            var expectedHeaders = "Header1: Value1\r\nHeader2: Value2";
            var propertyAccessorMock = new Mock<PropertyAccessor>();
            propertyAccessorMock
                .Setup(pa => pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E"))
                .Returns(expectedHeaders);

            var dyn = new DynamicIItem();
            dyn.Properties["PropertyAccessor"] = propertyAccessorMock.Object;

            var wrapper = CreateWrapper(dyn, _itemMock.Object, _eventsMock.Object, ImmutableHashSet.Create("IItem"));

            // Act
            var headers = wrapper.GetType()
                .GetMethod("GetRawHeaders", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(wrapper, null);

            // Assert
            Assert.AreEqual(expectedHeaders, headers);
        }

        [TestMethod]
        public void GetRawHeaders_ReturnsNull_WhenPropertyAccessorReturnsNull()
        {
            // Arrange
            var propertyAccessorMock = new Mock<PropertyAccessor>();
            propertyAccessorMock
                .Setup(pa => pa.GetProperty(It.IsAny<string>()))
                .Returns((string)null);

            var dyn = new DynamicIItem();
            dyn.Properties["PropertyAccessor"] = propertyAccessorMock.Object;

            var wrapper = CreateWrapper(dyn, _itemMock.Object, _eventsMock.Object, ImmutableHashSet.Create("IItem"));

            // Act
            var headers = wrapper.GetType()
                .GetMethod("GetRawHeaders", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(wrapper, null);

            // Assert
            Assert.IsNull(headers);
        }

        [TestMethod]
        public void GetRawHeaders_ReturnsNull_WhenExceptionIsThrown()
        {
            // Arrange
            var propertyAccessorMock = new Mock<PropertyAccessor>();
            propertyAccessorMock
                .Setup(pa => pa.GetProperty(It.IsAny<string>()))
                .Throws(new InvalidOperationException("Simulated exception"));
            _itemMock.Setup(m => m.PropertyAccessor).Returns(propertyAccessorMock.Object);
            LockMockObjects();
            

            // Act
            var headers = _wrapper.GetType()
                .GetMethod("GetRawHeaders", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(_wrapper, null);

            // Assert
            Assert.IsNull(headers);
        }

        [TestMethod]
        public void GetMessageId_ParsesMessageId_FromNormalizedHeaders()
        {
            // Arrange
            LockMockObjects();

            // Simulate folded Message-ID header with extra whitespace
            var rawNormalized = "From: test@example.com\r\nMessage-ID: <abc123@domain.com>\r\nTo: x@y.com";

            // Inject into _rawHeadersNormalized
            var normField = typeof(OutlookItemWrapper).GetField("_rawHeadersNormalized", BindingFlags.NonPublic | BindingFlags.Instance);
            normField.SetValue(_wrapper, new Lazy<string>(() => rawNormalized));

            var method = _wrapper.GetType().GetMethod("GetMessageId", BindingFlags.Instance | BindingFlags.NonPublic);

            // Act
            var messageId = method.Invoke(_wrapper, null);
            Console.WriteLine($"Expected Message-ID: abc123@domain.com");
            Console.WriteLine($"Actual   Message-ID: {messageId}");

            // Assert
            Assert.AreEqual("abc123@domain.com", messageId);
        }

        [TestMethod]
        public void GetMessageId_ParsesMessageId_FromNormalizedHeadersUnbracketed()
        {
            // Arrange
            LockMockObjects();

            // Simulate folded Message-ID header with extra whitespace
            var rawNormalized = "From: test@example.com\r\nMessage-ID: abc123@domain.com\r\nTo: x@y.com";

            // Inject into _rawHeadersNormalized
            var normField = typeof(OutlookItemWrapper).GetField("_rawHeadersNormalized", BindingFlags.NonPublic | BindingFlags.Instance);
            normField.SetValue(_wrapper, new Lazy<string>(() => rawNormalized));

            var method = _wrapper.GetType().GetMethod("GetMessageId", BindingFlags.Instance | BindingFlags.NonPublic);

            // Act
            var messageId = method.Invoke(_wrapper, null);
            Console.WriteLine($"Expected Message-ID: abc123@domain.com");
            Console.WriteLine($"Actual   Message-ID: {messageId}");

            // Assert
            Assert.AreEqual("abc123@domain.com", messageId);
        }


        [TestMethod]
        public void GetMessageId_ReturnsNull_WhenNoMessageIdInHeaders()
        {
            // Arrange
            var rawHeaders = "From: test@example.com\r\nTo: x@y.com";
            var dyn = new DynamicIItem();
            dyn.Properties["PropertyAccessor"] = null;

            var wrapper = CreateWrapper(dyn, _itemMock.Object, _eventsMock.Object, ImmutableHashSet.Create("IItem"));
            var rawHeadersField = typeof(OutlookItemWrapper).GetField("_rawHeaders", BindingFlags.NonPublic | BindingFlags.Instance);
            rawHeadersField.SetValue(wrapper, new Lazy<string>(() => rawHeaders));

            var getMessageId = wrapper.GetType().GetMethod("GetMessageId", BindingFlags.Instance | BindingFlags.NonPublic);

            // Act
            var messageId = getMessageId.Invoke(wrapper, null);

            // Assert
            Assert.IsNull(messageId);
        }

        [TestMethod]
        public void GetMessageId_ReturnsNull_WhenBracketIsUnclosed()
        {
            LockMockObjects();
            var malformed = "Message-ID: <abc@x.com";

            var field = typeof(OutlookItemWrapper).GetField("_rawHeadersNormalized", BindingFlags.NonPublic | BindingFlags.Instance);
            field.SetValue(_wrapper, new Lazy<string>(() => malformed));

            var method = _wrapper.GetType().GetMethod("GetMessageId", BindingFlags.Instance | BindingFlags.NonPublic);
            var result = method.Invoke(_wrapper, null);

            Assert.IsNull(result);
        }

        [TestMethod]
        public void GetMessageId_ReturnsNull_WhenBracketIsUnopened()
        {
            LockMockObjects();
            var malformed = "Message-ID: abc@x.com>";

            var field = typeof(OutlookItemWrapper).GetField("_rawHeadersNormalized", BindingFlags.NonPublic | BindingFlags.Instance);
            field.SetValue(_wrapper, new Lazy<string>(() => malformed));

            var method = _wrapper.GetType().GetMethod("GetMessageId", BindingFlags.Instance | BindingFlags.NonPublic);
            var result = method.Invoke(_wrapper, null);

            Assert.IsNull(result);
        }



        [TestMethod]
        public void NormalizeRawHeaders_UnfoldsAndSanitizesCorrectly()
        {
            // Arrange
            LockMockObjects();
            var folded = "From: test@example.com\r\nMessage-ID: <abc123@\r\n domain.com>\r\nTo: x@y.com\r\nX-Noise:\x01\x02value";

            // Inject into _rawHeaders
            var rawField = typeof(OutlookItemWrapper).GetField("_rawHeaders", BindingFlags.NonPublic | BindingFlags.Instance);
            rawField.SetValue(_wrapper, new Lazy<string>(() => folded));

            var method = _wrapper.GetType().GetMethod("NormalizeRawHeaders", BindingFlags.Instance | BindingFlags.NonPublic);

            // Act
            var normalized = (string)method.Invoke(_wrapper, null);

            // Assert
            string expected = "From: test@example.com\r\nMessage-ID: <abc123@ domain.com>\r\nTo: x@y.com\r\nX-Noise:value";
            Assert.AreEqual(expected, normalized);
        }

        [TestMethod]
        public void NormalizeRawHeaders_ReturnsNull_WhenHeadersAreWhitespace()
        {
            // Arrange
            LockMockObjects();
            var whitespaceHeaders = "   \r\n\t  ";
            var rawHeadersField = typeof(OutlookItemWrapper).GetField("_rawHeaders", BindingFlags.NonPublic | BindingFlags.Instance);
            rawHeadersField.SetValue(_wrapper, new Lazy<string>(() => whitespaceHeaders));

            // Act
            var method = _wrapper.GetType().GetMethod("NormalizeRawHeaders", BindingFlags.Instance | BindingFlags.NonPublic);
            var result = method.Invoke(_wrapper, null);

            // Assert
            Assert.IsNull(result);
        }

        [TestMethod]
        public void ReleaseComObject_CallsMarshalReleaseComObject_ForComObject()
        {
            // Arrange
            Type comType = Type.GetTypeFromProgID("Scripting.Dictionary");
            if (comType == null)
            {
                Assert.Inconclusive("Scripting.Dictionary COM type not available on this system.");
                return;
            }
            object comObj = Activator.CreateInstance(comType);
            var wrapper = CreateWrapper(_dynItem, _itemMock.Object, _eventsMock.Object, ImmutableHashSet.Create("IItem"));

            // Act & Assert
            // Should not throw
            try
            {
                var method = wrapper.GetType().GetMethod("ReleaseComObject", BindingFlags.Instance | BindingFlags.NonPublic);
                method.Invoke(wrapper, new object[] { comObj });
            }
            finally
            {
                Marshal.ReleaseComObject(comObj);
            }
        }

        [TestMethod]
        public void ReleaseComObject_DoesNothing_ForNull()
        {
            // Arrange
            var wrapper = CreateWrapper(_dynItem, _itemMock.Object, _eventsMock.Object, ImmutableHashSet.Create("IItem"));

            // Act & Assert
            var method = wrapper.GetType().GetMethod("ReleaseComObject", BindingFlags.Instance | BindingFlags.NonPublic);
            // Should not throw
            method.Invoke(wrapper, new object[] { null });
        }

        [TestMethod]
        public void MessageId_PropertySetGet_ReturnsValue()
        {
            // Arrange
            LockMockObjects();
            //var wrapper = CreateWrapper(_dynItem, _itemMock.Object, _eventsMock.Object, ImmutableHashSet.Create("IItem"));
            var expectedGet = "msgid-123";
            var expectedSet = "msgid-456";
            var messageIdField = typeof(OutlookItemWrapper).GetField("_messageId", BindingFlags.NonPublic | BindingFlags.Instance);
            messageIdField.SetValue(_wrapper, new Lazy<string>(() => expectedGet));

            // Act
            var actualGet = _wrapper.MessageId;
            (_wrapper as TestableOutlookItemWrapper).MessageId = expectedSet;
            var actualSet = _wrapper.MessageId;


            // Assert
            Console.WriteLine($"Expected Get: {expectedGet}");
            Console.WriteLine($"Actual Get:   {actualGet}");
            Console.WriteLine("");
            Console.WriteLine($"Expected Set: {expectedSet}");
            Console.WriteLine($"Actual Set:   {actualSet}");
            Assert.AreEqual(expectedGet, actualGet, "Get method did not function properly");
            Assert.AreEqual(expectedSet, actualSet, "Set method did not function properly");
        }

        [TestMethod]
        public void RawHeaders_PropertySetGet_ReturnsValue()
        {
            // Arrange
            //var wrapper = CreateWrapper(_dynItem, _itemMock.Object, _eventsMock.Object, ImmutableHashSet.Create("IItem"));
            LockMockObjects();
            var expectedGet = "Header1: Value1\r\nHeader2: Value2";
            var expectedSet = "Header1: NewValue1\r\nHeader2: NewValue2";
            var rawHeadersField = typeof(OutlookItemWrapper).GetField("_rawHeaders", BindingFlags.NonPublic | BindingFlags.Instance);
            rawHeadersField.SetValue(_wrapper, new Lazy<string>(() => expectedGet));

            // Act
            var actualGet = _wrapper.RawHeaders;
            (_wrapper as TestableOutlookItemWrapper).RawHeaders = expectedSet;
            var actualSet = _wrapper.RawHeaders;

            // Assert
            Console.WriteLine($"Expected Get: {expectedGet}");
            Console.WriteLine($"Actual Get:   {actualGet}");
            Console.WriteLine("");
            Console.WriteLine($"Expected Set: {expectedSet}");
            Console.WriteLine($"Actual Set:   {actualSet}");
            Assert.AreEqual(expectedGet, actualGet, "Get failed");
            Assert.AreEqual(expectedSet, actualSet, "Set failed");
        }

        [TestMethod]
        public void RawHeadersNormalized_PropertySetGet_ReturnsValue()
        {
            // Arrange            
            LockMockObjects();
            var expectedGet = "Header1: Value1\r\nHeader2: Value2";
            var expectedSet = "Header1: NewValue1\r\nHeader2: NewValue2";
            var rawHeadersField = typeof(OutlookItemWrapper).GetField("_rawHeadersNormalized", BindingFlags.NonPublic | BindingFlags.Instance);
            rawHeadersField.SetValue(_wrapper, new Lazy<string>(() => expectedGet));

            // Act
            var actualGet = _wrapper.RawHeadersNormalized;
            (_wrapper as TestableOutlookItemWrapper).RawHeadersNormalized = expectedSet;
            var actualSet = _wrapper.RawHeadersNormalized;

            // Assert
            Console.WriteLine($"Expected Get: {expectedGet}");
            Console.WriteLine($"Actual Get:   {actualGet}");
            Console.WriteLine("");
            Console.WriteLine($"Expected Set: {expectedSet}");
            Console.WriteLine($"Actual Set:   {actualSet}");
            Assert.AreEqual(expectedGet, actualGet, "Get failed");
            Assert.AreEqual(expectedSet, actualSet, "Set failed");
        }

        [TestMethod]
        public void InnerObject_Property_ReturnsUnderlyingItem()
        {
            // Arrange
            LockMockObjects();

            // Act & Assert
            Assert.AreEqual(_item, _wrapper.InnerObject);
        }

        [TestMethod]
        public void EqualsObject_ReturnsTrue_ForIdenticalReference()
        {
            // Arrange
            LockMockObjects();
            
            // Act
            var result = _wrapper.Equals((object)_wrapper);

            // Assert
            Assert.IsTrue(result);            
        }

        [TestMethod]
        public void EqualsObject_ReturnsTrue_ForEquivalentObjects()
        {
            // Arrange
            LockMockObjects();
            var otherMock = CreateIItemMock();
            var otherTypes = ImmutableHashSet.Create(otherMock.Object.GetType().Name);
            var otherWrapper = CreateWrapper(otherMock.Object, _events, otherTypes);

            // Use a custom comparer that returns true for equality
            var customComparer = new Mock<IEqualityComparer<IItem>>();
            customComparer.Setup(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>())).Returns(true);
            _wrapper.EqualityComparer = customComparer.Object;

            // Act
            var result = _wrapper.Equals((object)otherWrapper);

            // Assert
            Assert.IsTrue(result);
            customComparer.Verify(c => c.Equals(_wrapper, otherWrapper), Times.Once);
        }

        [TestMethod]
        public void EqualsObject_ReturnsFalse_ForDifferentObjects()
        {
            // Arrange
            LockMockObjects();
            var otherObject = new object(); // Not an IItem

            // Use a custom comparer that returns false for equality
            var customComparer = new Mock<IEqualityComparer<IItem>>();
            customComparer.Setup(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>())).Returns(false);
            _wrapper.EqualityComparer = customComparer.Object;

            // Act
            var result = _wrapper.Equals((object)otherObject);

            // Assert
            Assert.IsFalse(result);
            customComparer.Verify(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>()), Times.Never);
        }        
    }
}