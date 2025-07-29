using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using FluentAssertions;
using Microsoft.Office.Interop.Outlook;
using Gsync.OutlookInterop.Item;
using System.Reflection;
using System.Dynamic;

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
            var dynField = typeof(OutlookItemWrapper).GetField("_dyn", BindingFlags.NonPublic | BindingFlags.Instance);
            dynField.SetValue(_wrapper, _dynMailItem);
        }

        [TestMethod]
        public void BCC_GetSet_ShouldForwardToMailItem()
        {
            _wrapper.BCC = "abc";
            _wrapper.BCC.Should().Be("abc");
        }

        [TestMethod]
        public void ReminderSet_GetSet_ShouldForwardToMailItem()
        {
            _wrapper.ReminderSet = true;
            _wrapper.ReminderSet.Should().BeTrue();
        }

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
            _dynMailItem.Methods["Forward"] = new Func<MailItem>(() => expectedMailItem);

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
        public void Dispose_ShouldDetachMailItemEventsAndDisposeBase()
        {
            var wrapperType = typeof(MailItemWrapper);
            var mailItemEventsAttachedField = wrapperType.GetField("_mailItemEventsAttached", BindingFlags.NonPublic | BindingFlags.Instance);
            mailItemEventsAttachedField.SetValue(_wrapper, true);

            _wrapper.Dispose();

            ((bool)mailItemEventsAttachedField.GetValue(_wrapper)).Should().BeFalse();
        }

        [TestMethod]
        public void CustomActionEvent_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.CustomAction += (a, r, ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnCustomAction", BindingFlags.NonPublic | BindingFlags.Instance);
            object[] parameters = { new object(), new object(), false };
            method.Invoke(_wrapper, parameters);
            eventCalled.Should().BeTrue();
        }

        [TestMethod]
        public void ReplyEvent_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.ReplyEvent += (r, ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnReply", BindingFlags.NonPublic | BindingFlags.Instance);
            object[] parameters = { new object(), false };
            method.Invoke(_wrapper, parameters);
            eventCalled.Should().BeTrue();
        }

        [TestMethod]
        public void SendEvent_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.SendEvent += (ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnSend", BindingFlags.NonPublic | BindingFlags.Instance);
            object[] parameters = { false };
            method.Invoke(_wrapper, parameters);
            eventCalled.Should().BeTrue();
        }

        [TestMethod]
        public void BeforeCheckNames_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            _wrapper.BeforeCheckNames += (ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnBeforeCheckNames", BindingFlags.NonPublic | BindingFlags.Instance);
            object[] parameters = { false };
            method.Invoke(_wrapper, parameters);
            eventCalled.Should().BeTrue();
        }

        [TestMethod]
        public void BeforeAttachmentSave_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            var attachment = new Mock<Attachment>().Object;
            _wrapper.BeforeAttachmentSave += (a, ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnBeforeAttachmentSave", BindingFlags.NonPublic | BindingFlags.Instance);
            object[] parameters = { attachment, false };
            method.Invoke(_wrapper, parameters);
            eventCalled.Should().BeTrue();
        }

        [TestMethod]
        public void BeforeAttachmentAdd_ShouldInvokeHandler()
        {
            bool eventCalled = false;
            var attachment = new Mock<Attachment>().Object;
            _wrapper.BeforeAttachmentAdd += (a, ref c) => eventCalled = true;
            var method = typeof(MailItemWrapper).GetMethod("OnBeforeAttachmentAdd", BindingFlags.NonPublic | BindingFlags.Instance);
            object[] parameters = { attachment, false };
            method.Invoke(_wrapper, parameters);
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
            object[] parameters = { false };
            method.Invoke(_wrapper, parameters);
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
    }
}
