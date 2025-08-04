using FluentAssertions;
using Gsync.OutlookInterop.Interfaces.Items;
using Gsync.OutlookInterop.Item;
using Gsync.Utilities.HelperClasses;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Reflection;

namespace Gsync.Test.OutlookInterop.Item
{
    [TestClass]
    public class MailItemWrapperTests
    {
        private Mock<MailItem> _mailItemMock;
        private Mock<ItemEvents_10_Event> _eventsMock;
        private MailItemWrapper _wrapper;
        private DynamicMailItem _dynMailItem;

        [TestInitialize]
        public void Setup()
        {
            Console.SetOut(new DebugTextWriter());
            _mailItemMock = new Mock<MailItem>();
            _eventsMock = new Mock<ItemEvents_10_Event>();
            _wrapper = (MailItemWrapper)Activator.CreateInstance(
                typeof(MailItemWrapper),
                BindingFlags.NonPublic | BindingFlags.Instance,
                null,
                new object[] { _mailItemMock.Object, _eventsMock.Object },
                null
            );
            _dynMailItem = new DynamicMailItem();

            // Inject _dynMailItem into private _dyn field
            var dynField = typeof(MailItemWrapper).GetField("_dyn", BindingFlags.NonPublic | BindingFlags.Instance);
            dynField.SetValue(_wrapper, _dynMailItem);
        }

        private Mock<MailItem> CreateMailItemMock()
        {
            var mock = new Mock<MailItem>();
            mock.SetupAllProperties();
            return mock;
        }

        // Property forwarding tests
        [TestMethod]
        public void BCC_GetSet_ShouldWorkThroughWrapper()
        {
            _dynMailItem.Properties["BCC"] = "mybcc";
            _wrapper.BCC.Should().Be("mybcc");
            _wrapper.BCC = "newbcc";
            _dynMailItem.Properties["BCC"].Should().Be("newbcc");
        }
        [TestMethod]
        public void ReminderSet_GetSet_ShouldWorkThroughWrapper()
        {
            _dynMailItem.Properties["ReminderSet"] = false;
            _wrapper.ReminderSet.Should().BeFalse();
            _wrapper.ReminderSet = true;
            ((bool)_dynMailItem.Properties["ReminderSet"]).Should().BeTrue();
        }
        [TestMethod] public void CC_GetSet_ShouldForward() { _wrapper.CC = "foo"; _wrapper.CC.Should().Be("foo"); }
        [TestMethod] public void DeferredDeliveryTime_GetSet_ShouldForward() { _wrapper.DeferredDeliveryTime = "2025-01-01"; _wrapper.DeferredDeliveryTime.Should().Be("2025-01-01"); }
        [TestMethod] public void DeleteAfterSubmit_GetSet_ShouldForward() { _wrapper.DeleteAfterSubmit = "bar"; _wrapper.DeleteAfterSubmit.Should().Be("bar"); }
        [TestMethod] public void FlagRequest_GetSet_ShouldForward() { _wrapper.FlagRequest = "Urgent"; _wrapper.FlagRequest.Should().Be("Urgent"); }
        [TestMethod] public void ReceivedByName_Get_ShouldForward() { _dynMailItem.Properties["ReceivedByName"] = "me"; _wrapper.ReceivedByName.Should().Be("me"); }
        [TestMethod] public void ReceivedOnBehalfOfName_Get_ShouldForward() { _dynMailItem.Properties["ReceivedOnBehalfOfName"] = "her"; _wrapper.ReceivedOnBehalfOfName.Should().Be("her"); }
        [TestMethod] public void ReceivedTime_Get_ShouldForward() { var now = DateTime.Now; _dynMailItem.Properties["ReceivedTime"] = now; _wrapper.ReceivedTime.Should().Be(now); }
        [TestMethod] public void RecipientReassignmentProhibited_GetSet_ShouldForward() { _wrapper.RecipientReassignmentProhibited = "true"; _wrapper.RecipientReassignmentProhibited.Should().Be("true"); }
        [TestMethod] public void ReminderOverrideDefault_GetSet_ShouldForward() { _wrapper.ReminderOverrideDefault = true; _wrapper.ReminderOverrideDefault.Should().BeTrue(); }
        [TestMethod] public void ReminderPlaySound_GetSet_ShouldForward() { _wrapper.ReminderPlaySound = true; _wrapper.ReminderPlaySound.Should().BeTrue(); }
        [TestMethod] public void ReminderSoundFile_GetSet_ShouldForward() { _wrapper.ReminderSoundFile = "file.wav"; _wrapper.ReminderSoundFile.Should().Be("file.wav"); }
        [TestMethod] public void ReminderTime_GetSet_ShouldForward() { var dt = DateTime.Today; _wrapper.ReminderTime = dt; _wrapper.ReminderTime.Should().Be(dt); }
        [TestMethod] public void ReplyRecipientNames_Get_ShouldForward() { _dynMailItem.Properties["ReplyRecipientNames"] = "a;b"; _wrapper.ReplyRecipientNames.Should().Be("a;b"); }
        [TestMethod] public void SaveSentMessageFolder_GetSet_ShouldForward() { _wrapper.SaveSentMessageFolder = 99; _wrapper.SaveSentMessageFolder.Should().Be(99); }
        [TestMethod] public void SenderEmailType_Get_ShouldForward() { _dynMailItem.Properties["SenderEmailType"] = "SMTP"; _wrapper.SenderEmailType.Should().Be("SMTP"); }
        [TestMethod] public void SentOnBehalfOfName_GetSet_ShouldForward() { _wrapper.SentOnBehalfOfName = "John"; _wrapper.SentOnBehalfOfName.Should().Be("John"); }
        [TestMethod] public void SentOn_Get_ShouldForward() { var sent = DateTime.Today; _dynMailItem.Properties["SentOn"] = sent; _wrapper.SentOn.Should().Be(sent); }
        [TestMethod] public void Submitted_Get_ShouldForward() { _dynMailItem.Properties["Submitted"] = true; _wrapper.Submitted.Should().BeTrue(); }
        [TestMethod] public void To_GetSet_ShouldForward() { _wrapper.To = "foo@bar.com"; _wrapper.To.Should().Be("foo@bar.com"); }
        [TestMethod] public void VotingOptions_GetSet_ShouldForward() { _wrapper.VotingOptions = "Yes;No"; _wrapper.VotingOptions.Should().Be("Yes;No"); }
        [TestMethod] public void VotingResponse_GetSet_ShouldForward() { _wrapper.VotingResponse = "Accepted"; _wrapper.VotingResponse.Should().Be("Accepted"); }

        [TestMethod]
        public void Property_Getter_MissingProperty_ShouldThrowRuntimeBinderException()
        {
            System.Action act = () => { var _ = _wrapper.FlagRequest; };
            act.Should().Throw<Microsoft.CSharp.RuntimeBinder.RuntimeBinderException>();
        }

        // Method forwarding tests
        [TestMethod]
        public void Send_ShouldInvokeMailItemSend()
        {
            bool sendCalled = false;
            _dynMailItem.Methods["Send"] = new System.Action(() => sendCalled = true);

            _wrapper.Send();
            sendCalled.Should().BeTrue();
        }
        [TestMethod]
        public void Forward_ShouldInvokeMailItemForward()
        {
            var expectedMailItem = new Mock<MailItem>().Object;
            _dynMailItem.Methods["Forward"] = new System.Func<MailItem>(() => expectedMailItem);

            var result = _wrapper.Forward();
            result.Should().Be(expectedMailItem);
        }
        [TestMethod]
        public void ClearConversationIndex_ShouldInvokeMailItemMethod()
        {
            bool called = false;
            _dynMailItem.Methods["ClearConversationIndex"] = new System.Action(() => called = true);

            _wrapper.ClearConversationIndex();
            called.Should().BeTrue();
        }
        [TestMethod]
        public void Reply_ShouldInvokeMailItemReply()
        {
            var expectedMailItem = new Mock<MailItem>().Object;
            _dynMailItem.Methods["Reply"] = new System.Func<MailItem>(() => expectedMailItem);
            _wrapper.Reply().Should().Be(expectedMailItem);
        }
        [TestMethod]
        public void ReplyAll_ShouldInvokeMailItemReplyAll()
        {
            var expectedMailItem = new Mock<MailItem>().Object;
            _dynMailItem.Methods["ReplyAll"] = new System.Func<MailItem>(() => expectedMailItem);
            _wrapper.ReplyAll().Should().Be(expectedMailItem);
        }
        [TestMethod]
        public void ImportanceChanged_ShouldInvokeMailItemMethod()
        {
            bool called = false;
            _dynMailItem.Methods["ImportanceChanged"] = new System.Action(() => called = true);
            _wrapper.ImportanceChanged();
            called.Should().BeTrue();
        }
        [TestMethod]
        public void Send_WhenMethodThrows_ShouldPropagateException()
        {
            _dynMailItem.Methods["Send"] = new System.Action(() => throw new InvalidOperationException("Test"));

            System.Action act = () => _wrapper.Send();

            var ex = act.Should()
                .Throw<TargetInvocationException>()
                .Which;

            ex.InnerException.Should().BeOfType<InvalidOperationException>();
            ex.InnerException.Message.Should().Be("Test");
        }


        // Event bridging coverage
        [TestMethod]
        public void CustomActionEvent_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.CustomAction += (a, r, ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnCustomAction", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, new object[] { new object(), new object(), false });
            eventCalled.Should().BeTrue();
        }
        [TestMethod]
        public void ReplyEvent_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.ReplyEvent += (r, ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnReply", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, new object[] { new object(), false });
            eventCalled.Should().BeTrue();
        }
        [TestMethod]
        public void SendEvent_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.SendEvent += (ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnSend", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, new object[] { false });
            eventCalled.Should().BeTrue();
        }
        [TestMethod]
        public void BeforeCheckNames_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.BeforeCheckNames += (ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnBeforeCheckNames", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, new object[] { false });
            eventCalled.Should().BeTrue();
        }
        [TestMethod]
        public void BeforeAttachmentSave_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            var attachment = new Mock<Attachment>().Object;
            _wrapper.BeforeAttachmentSave += (a, ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnBeforeAttachmentSave", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, new object[] { attachment, false });
            eventCalled.Should().BeTrue();
        }
        [TestMethod]
        public void BeforeAttachmentAdd_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            var attachment = new Mock<Attachment>().Object;
            _wrapper.BeforeAttachmentAdd += (a, ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnBeforeAttachmentAdd", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, new object[] { attachment, false });
            eventCalled.Should().BeTrue();
        }
        [TestMethod]
        public void Unload_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.Unload += () => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnUnload", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, null);
            eventCalled.Should().BeTrue();
        }
        [TestMethod]
        public void BeforeAutoSave_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.BeforeAutoSave += (ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnBeforeAutoSave", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, new object[] { false });
            eventCalled.Should().BeTrue();
        }
        [TestMethod]
        public void BeforeRead_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.BeforeRead += () => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnBeforeRead", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, null);
            eventCalled.Should().BeTrue();
        }

        // Constructor/static method coverage

        [TestMethod]
        public void MailItemWrapper_ObjectCtor_ShouldSetMailItemAndAttachEvents()
        {
            var mailItemMock = new Mock<MailItem>().Object;
            var wrapper = new TestableMailItemWrapper(mailItemMock);

            var mailItemField = typeof(MailItemWrapper).GetField("_mailItem", BindingFlags.NonPublic | BindingFlags.Instance);
            mailItemField.GetValue(wrapper).Should().Be(mailItemMock);
            ((TestableMailItemWrapper)wrapper).AttachMailItemEventsCalled.Should().BeTrue();
        }
        [TestMethod]
        public void MailItemWrapper_ObjectCtor_ShouldThrowIfNotMailItem()
        {
            var notMailItem = new object();
            var act = () => new TestableMailItemWrapper(notMailItem);
            act.Should().Throw<ArgumentException>();
        }
        [TestMethod]
        public void MailItemWrapper_ProtectedCtor_WithSupportedTypes_ShouldSetMailItemAndAttachEvents()
        {
            var mailItemMock = new Mock<MailItem>().Object;
            var eventsMock = new Mock<ItemEvents_10_Event>().Object;
            var supportedTypes = ImmutableHashSet.Create("MailItem");

            var wrapper = new TestableMailItemWrapper(mailItemMock, eventsMock, supportedTypes);

            var mailItemField = typeof(MailItemWrapper).GetField("_mailItem", BindingFlags.NonPublic | BindingFlags.Instance);
            mailItemField.GetValue(wrapper).Should().Be(mailItemMock);
            wrapper.AttachMailItemEventsCalled.Should().BeTrue();
        }
        [TestMethod]
        public void MailItemWrapper_ProtectedCtor_WithSupportedTypes_ShouldThrowIfNotMailItem()
        {
            var notMailItem = new object();
            var eventsMock = new Mock<ItemEvents_10_Event>().Object;
            var supportedTypes = ImmutableHashSet.Create("MailItem");

            var act = () => new TestableMailItemWrapper(notMailItem, eventsMock, supportedTypes);
            act.Should().Throw<ArgumentException>();
        }
        [TestMethod]
        public void MailItemWrapper_BaseWrapperCtor_ShouldSetMailItemAndAttachEvents()
        {
            var mailItemMock = new Mock<MailItem>().Object;
            var eventsMock = new Mock<ItemEvents_10_Event>().Object;
            var supportedTypes = ImmutableHashSet.Create("MailItem");            
            var baseWrapper = new StubMailItemWrapper(mailItemMock, eventsMock, supportedTypes);

            var wrapper = new TestableMailItemWrapper(baseWrapper);
            Console.WriteLine(string.Join(",", baseWrapper.SupportedTypes)); // Debug

            var mailItemField = typeof(MailItemWrapper).GetField("_mailItem", BindingFlags.NonPublic | BindingFlags.Instance);
            mailItemField.GetValue(wrapper).Should().Be(mailItemMock);
            wrapper.AttachMailItemEventsCalled.Should().BeTrue();
        }
        [TestMethod]
        public void MailItemWrapper_BaseWrapperCtor_ShouldThrowIfNotMailItem()
        {
            //var notMailItem = new object();
            var notMailItem = new Mock<MeetingItem>().Object; // MeetingItem is not a MailItem
            //var eventsMock = new Mock<ItemEvents_10_Event>().Object;
            //var supportedTypes = ImmutableHashSet.Create(notMailItem.GetType().Name);

            //var baseWrapper = new StubOutlookItemWrapper(notMailItem, eventsMock, supportedTypes);
            //var baseWrapper = new Mock<OutlookItemWrapper>(notMailItem, eventsMock, supportedTypes) { CallBase = true }.Object;
            //var baseWrapper = new Mock<OutlookItemWrapper>(notMailItem) { CallBase = true }.Object;
            //typeof(OutlookItemWrapper).GetField("_item", BindingFlags.NonPublic | BindingFlags.Instance)
            //    .SetValue(baseWrapper, notMailItem);
            var baseWrapper = new OutlookItemWrapper(notMailItem);

            var act = () => new MailItemWrapper(baseWrapper);
            //var act = () => new TestableMailItemWrapper(baseWrapper);
            var e = act.Should().Throw<ArgumentNullException>().Which;
            Console.WriteLine($"The correct {nameof(ArgumentNullException)} was thrown " +
                $"with the following message: {e.Message}");            
        }
        [TestMethod]
        public void GetCom10Events_ShouldReturnComEventsField()
        {
            var mailItemMock = new Mock<MailItem>().Object;
            var eventsMock = new Mock<ItemEvents_10_Event>().Object;
            var supportedTypes = ImmutableHashSet.Create("MailItem");

            var baseWrapper = new Mock<MailItemWrapper>(mailItemMock, eventsMock, supportedTypes) { CallBase = true }.Object;
            typeof(MailItemWrapper).GetField("_comEvents", BindingFlags.NonPublic | BindingFlags.Instance)
                .SetValue(baseWrapper, eventsMock);

            var result = TestableMailItemWrapper.GetCom10Events(baseWrapper);

            result.Should().Be(eventsMock);
        }
        [TestMethod]
        public void Dispose_ShouldDetachMailItemEventsAndDisposeBase()
        {
            var wrapperType = typeof(MailItemWrapper);
            var mailItemEventsAttachedField = wrapperType.GetField("_mailItemEventsAttached", BindingFlags.NonPublic | BindingFlags.Instance);
            mailItemEventsAttachedField.SetValue(_wrapper, true);

            _wrapper.Dispose();

            ((bool)mailItemEventsAttachedField.GetValue(_wrapper)).Should().BeFalse();
        }

        [TestMethod]
        public void EqualityComparer_Default_IsMailItemEqualityComparer()
        {
            var mock = CreateMailItemMock();
            var wrapper = new MailItemWrapper(mock.Object);
            Assert.IsInstanceOfType(wrapper.EqualityComparer, typeof(MailItemEqualityComparer));
        }

        [TestMethod]
        public void EqualityComparer_CanBeInjected_AndUsed()
        {
            var mock = CreateMailItemMock();
            var customComparer = new Mock<IEqualityComparer<IMailItem>>();
            customComparer.Setup(c => c.Equals(It.IsAny<IMailItem>(), It.IsAny<IMailItem>())).Returns(true);
            customComparer.Setup(c => c.GetHashCode(It.IsAny<IMailItem>())).Returns(123);

            var wrapper = new MailItemWrapper(mock.Object)
            {
                EqualityComparer = customComparer.Object
            };

            Assert.AreEqual(customComparer.Object, wrapper.EqualityComparer);

            // Should use the injected comparer
            Assert.IsTrue(wrapper.Equals((IMailItem)null));
            Assert.AreEqual(123, wrapper.GetHashCode());
        }

        [TestMethod]
        public void EqualsIItem_UsesEqualityComparer()
        {
            var mock = CreateMailItemMock();
            var otherMock = CreateMailItemMock();
            var customComparer = new Mock<IEqualityComparer<IMailItem>>();
            customComparer.Setup(c => c.Equals(It.IsAny<IMailItem>(), It.IsAny<IMailItem>())).Returns(false);
            var wrapper = new MailItemWrapper(mock.Object)
            {
                EqualityComparer = customComparer.Object
            };
            var otherWrapper = new MailItemWrapper(otherMock.Object);

            Assert.IsFalse(wrapper.Equals(otherWrapper));
            customComparer.Verify(c => c.Equals(It.IsAny<IMailItem>(), It.IsAny<IMailItem>()), Times.Once);
        }

        [TestMethod]
        public void EqualsObject_UsesEqualityComparer()
        {
            var mock = CreateMailItemMock();
            var otherMock = CreateMailItemMock();
            var customComparer = new Mock<IEqualityComparer<IItem>>();
            customComparer.Setup(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>())).Returns(true);

            var wrapper = new MailItemWrapper(mock.Object)
            {
                EqualityComparer = customComparer.Object
            };
            var otherWrapper = new MailItemWrapper(otherMock.Object);

            Assert.IsTrue(wrapper.Equals((object)otherWrapper));
            customComparer.Verify(c => c.Equals(It.IsAny<IItem>(), It.IsAny<IItem>()), Times.Once);
        }

        [TestMethod]
        public void GetHashCode_UsesEqualityComparer()
        {
            var mock = CreateMailItemMock();
            var customComparer = new Mock<IEqualityComparer<IItem>>();
            customComparer.Setup(c => c.GetHashCode(It.IsAny<IItem>())).Returns(42);

            var wrapper = new MailItemWrapper(mock.Object)
            {
                EqualityComparer = customComparer.Object
            };

            Assert.AreEqual(42, wrapper.GetHashCode());
            customComparer.Verify(c => c.GetHashCode(wrapper), Times.Once);
        }

        [TestMethod]
        public void ConversationID_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.ConversationID).Returns("ConvId");
            var wrapper = new MailItemWrapper(mock.Object);
            Assert.AreEqual("ConvId", wrapper.ConversationID);
        }

        [TestMethod]
        public void HTMLBody_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.HTMLBody, "TestHtml");
            var wrapper = new MailItemWrapper(mock.Object);
            wrapper.HTMLBody = "NewHtml";
            Assert.AreEqual("NewHtml", wrapper.HTMLBody);
        }
        [TestMethod]
        public void SenderEmailAddress_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.SenderEmailAddress).Returns("test@example.com");
            var wrapper = new MailItemWrapper(mock.Object);
            Assert.AreEqual("test@example.com", wrapper.SenderEmailAddress);
        }
        [TestMethod]
        public void SenderName_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.SenderName).Returns("Sender");
            var wrapper = new MailItemWrapper(mock.Object);
            Assert.AreEqual("Sender", wrapper.SenderName);
        }

        [TestMethod]
        public void ShowCategoriesDialog_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.ShowCategoriesDialog()).Verifiable();
            var wrapper = new MailItemWrapper(mock.Object);
            wrapper.ShowCategoriesDialog();
            mock.Verify(m => m.ShowCategoriesDialog(), Times.Once);
        }

        [TestMethod]
        public void ItemProperties_Property_ReturnsValue()
        {
            // Arrange
            var itemProperties = new Mock<ItemProperties>().Object;
            _dynMailItem.Properties["ItemProperties"] = itemProperties;

            // Act & Assert
            _wrapper.ItemProperties.Should().Be(itemProperties);
        }

        [TestMethod]
        public void Recipients_Property_ReturnsValue()
        {
            // Arrange
            var recipients = new Mock<Recipients>().Object;
            _dynMailItem.Properties["Recipients"] = recipients;

            // Act & Assert
            _wrapper.Recipients.Should().Be(recipients);
        }

        [TestMethod]
        public void ReplyRecipients_Property_ReturnsValue()
        {
            // Arrange
            var replyRecipients = new Mock<Recipients>().Object;
            _dynMailItem.Properties["ReplyRecipients"] = replyRecipients;

            // Act & Assert
            _wrapper.ReplyRecipients.Should().Be(replyRecipients);
        }

        [TestMethod]
        public void OnCustomPropertyChange_RaisesEvent()
        {
            // Arrange
            bool called = false;
            string expectedName = "TestProp";
            _wrapper.CustomPropertyChange += name =>
            {
                called = true;
                name.Should().Be(expectedName);
            };

            // Act
            var method = typeof(MailItemWrapper).GetMethod("OnCustomPropertyChange", BindingFlags.NonPublic | BindingFlags.Instance);
            method.Invoke(_wrapper, new object[] { expectedName });

            // Assert
            called.Should().BeTrue();
        }

        [TestMethod]
        public void OnForward_RaisesEvent()
        {
            // Arrange
            bool called = false;
            object forwardObj = new object();
            bool cancel = false;
            _wrapper.ForwardEvent += (object obj, ref bool c) =>
            {
                called = true;
                obj.Should().Be(forwardObj);
                c = true;
            };

            // Act
            var method = typeof(MailItemWrapper).GetMethod("OnForward", BindingFlags.NonPublic | BindingFlags.Instance);
            object[] parameters = new object[] { forwardObj, cancel };
            method.Invoke(_wrapper, parameters);

            // Assert
            called.Should().BeTrue();
        }

        [TestMethod]
        public void OnReplyAll_RaisesEvent()
        {
            // Arrange
            bool called = false;
            object responseObj = new object();
            bool cancel = false;
            _wrapper.ReplyAllEvent += (object obj, ref bool c) =>
            {
                called = true;
                obj.Should().Be(responseObj);
                c = true;
            };

            // Act
            var method = typeof(MailItemWrapper).GetMethod("OnReplyAll", BindingFlags.NonPublic | BindingFlags.Instance);
            object[] parameters = new object[] { responseObj, cancel };
            method.Invoke(_wrapper, parameters);

            // Assert
            called.Should().BeTrue();
        }

        [TestMethod]
        public void EqualsObject_ReturnsTrue_ForIdenticalReference()
        {            
            // Act
            var result = _wrapper.Equals((object)_wrapper);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EqualsObject_ReturnsTrue_ForEquivalentObjects()
        {
            // Arrange            
            var otherMock = CreateMailItemMock();
            var otherTypes = ImmutableHashSet.Create(otherMock.Object.GetType().Name);
            //var otherWrapper = new TestableMailItemWrapper(otherMock.Object, _eventsMock.Object, otherTypes);
            var otherWrapper = new MailItemWrapper(otherMock.Object);

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
