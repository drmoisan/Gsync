using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using Gsync.OutlookInterop.Item;
using Gsync.OutlookInterop.Interfaces.Items;
using Newtonsoft.Json;

namespace Gsync.Test.OutlookInterop.Item
{
    [TestClass]
    public class DetachedOutlookItemTests
    {
        private Mock<IItem> CreateMockItem()
        {
            var mock = new Mock<IItem>();
            mock.SetupAllProperties();
            mock.SetupGet(m => m.BillingInformation).Returns("BillInfo");
            mock.SetupGet(m => m.Body).Returns("BodyValue");
            mock.SetupGet(m => m.Categories).Returns("Cats");
            mock.SetupGet(m => m.Class).Returns(OlObjectClass.olMail);
            mock.SetupGet(m => m.Companies).Returns("MyCo");
            mock.SetupGet(m => m.ConversationID).Returns("ConvId");
            var now = DateTime.Now;
            mock.SetupGet(m => m.CreationTime).Returns(now);
            mock.SetupGet(m => m.EntryID).Returns("itemEntryId");
            mock.SetupGet(m => m.HTMLBody).Returns("HtmlBody");
            mock.SetupGet(m => m.Importance).Returns(OlImportance.olImportanceHigh);
            mock.SetupGet(m => m.LastModificationTime).Returns(now);
            mock.SetupGet(m => m.MessageClass).Returns("IPM.Note");
            mock.SetupGet(m => m.Mileage).Returns("100mi");
            mock.SetupGet(m => m.NoAging).Returns(true);
            mock.SetupGet(m => m.OutlookInternalVersion).Returns(42);
            mock.SetupGet(m => m.OutlookVersion).Returns("16.0");
            mock.SetupGet(m => m.Saved).Returns(true);
            mock.SetupGet(m => m.SenderEmailAddress).Returns("foo@bar.com");
            mock.SetupGet(m => m.SenderName).Returns("FooBar");
            mock.SetupGet(m => m.Sensitivity).Returns(OlSensitivity.olPrivate);
            mock.SetupGet(m => m.Size).Returns(1234);
            mock.SetupGet(m => m.Subject).Returns("Test Subject");
            mock.SetupGet(m => m.UnRead).Returns(true);
            // Parent as MAPIFolder
            var folderMock = new Mock<MAPIFolder>();
            folderMock.SetupGet(f => f.StoreID).Returns("storeId123");
            folderMock.SetupGet(f => f.EntryID).Returns("folderEntryId456");
            mock.SetupGet(m => m.Parent).Returns(folderMock.Object);

            return mock;
        }

        [TestMethod]
        public void Constructor_Parameterless_NotNull()
        {
            var detached = new DetachedOutlookItem();
            Assert.IsNotNull(detached);
        }

        [TestMethod]
        public void Constructor_CopiesValueProperties()
        {
            var mock = CreateMockItem();
            var detached = new DetachedOutlookItem(mock.Object);

            Assert.AreEqual("BillInfo", detached.BillingInformation);
            Assert.AreEqual("BodyValue", detached.Body);
            Assert.AreEqual("Cats", detached.Categories);
            Assert.AreEqual(OlObjectClass.olMail, detached.Class);
            Assert.AreEqual("MyCo", detached.Companies);
            Assert.AreEqual("ConvId", detached.ConversationID);
            Assert.AreEqual(mock.Object.CreationTime, detached.CreationTime);
            Assert.AreEqual("itemEntryId", detached.EntryID);
            Assert.AreEqual("HtmlBody", detached.HTMLBody);
            Assert.AreEqual(OlImportance.olImportanceHigh, detached.Importance);
            Assert.AreEqual(mock.Object.LastModificationTime, detached.LastModificationTime);
            Assert.AreEqual("IPM.Note", detached.MessageClass);
            Assert.AreEqual("100mi", detached.Mileage);
            Assert.AreEqual(true, detached.NoAging);
            Assert.AreEqual(42, detached.OutlookInternalVersion);
            Assert.AreEqual("16.0", detached.OutlookVersion);
            Assert.AreEqual(true, detached.Saved);
            Assert.AreEqual("foo@bar.com", detached.SenderEmailAddress);
            Assert.AreEqual("FooBar", detached.SenderName);
            Assert.AreEqual(OlSensitivity.olPrivate, detached.Sensitivity);
            Assert.AreEqual(1234, detached.Size);
            Assert.AreEqual("Test Subject", detached.Subject);
            Assert.AreEqual(true, detached.UnRead);
            Assert.AreEqual("storeId123", detached.StoreID);
            Assert.AreEqual("folderEntryId456", detached.ParentFolderEntryID);
        }

        [TestMethod]
        public void Constructor_WhenItemIsNull_ThrowsArgumentNullException()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new DetachedOutlookItem(null));
        }

        [TestMethod]
        public void ComProperties_AreAlwaysNull()
        {
            var detached = new DetachedOutlookItem(CreateMockItem().Object);
            Assert.IsNull(detached.Application);
            Assert.IsNull(detached.Attachments);
            Assert.IsNull(detached.ItemProperties);
            Assert.IsNull(detached.Session);
            Assert.IsNull(detached.InnerObject);
            Assert.IsNull(detached.Parent);
        }

        [TestMethod]
        public void Methods_Throw_NotSupportedException()
        {
            var detached = new DetachedOutlookItem(CreateMockItem().Object);
            Assert.ThrowsException<NotSupportedException>(() => detached.Close(OlInspectorClose.olSave));
            Assert.ThrowsException<NotSupportedException>(() => detached.Copy());
            Assert.ThrowsException<NotSupportedException>(() => detached.Delete());
            Assert.ThrowsException<NotSupportedException>(() => detached.Display());
            Assert.ThrowsException<NotSupportedException>(() => detached.Move(null));
            Assert.ThrowsException<NotSupportedException>(() => detached.PrintOut());
            Assert.ThrowsException<NotSupportedException>(() => detached.Save());
            Assert.ThrowsException<NotSupportedException>(() => detached.SaveAs("x"));
            Assert.ThrowsException<NotSupportedException>(() => detached.ShowCategoriesDialog());
        }

        [TestMethod]
        public void Events_SubscribeAndUnsubscribe_NoError()
        {
            var detached = new DetachedOutlookItem(CreateMockItem().Object);

            IItem.AttachmentAddEventHandler attachmentAddHandler = (Attachment a) => { };
            IItem.AttachmentReadEventHandler attachmentReadHandler = (Attachment a) => { };
            IItem.AttachmentRemoveEventHandler attachmentRemoveHandler = (Attachment a) => { };
            IItem.BeforeDeleteEventHandler beforeDeleteHandler = (object obj, ref bool c) => { };
            IItem.CloseEventHandler closeEventHandler = (ref bool c) => { };
            IItem.OpenEventHandler openEventHandler = (ref bool c) => { };
            IItem.PropertyChangeEventHandler propertyChangeHandler = (string s) => { };
            IItem.ReadEventHandler readHandler = () => { };
            IItem.WriteEventHandler writeEventHandler = (ref bool c) => { };

            // Subscribe and unsubscribe each event; should not throw
            detached.AttachmentAdd += attachmentAddHandler;
            detached.AttachmentAdd -= attachmentAddHandler;

            detached.AttachmentRead += attachmentReadHandler;
            detached.AttachmentRead -= attachmentReadHandler;

            detached.AttachmentRemove += attachmentRemoveHandler;
            detached.AttachmentRemove -= attachmentRemoveHandler;

            detached.BeforeDelete += beforeDeleteHandler;
            detached.BeforeDelete -= beforeDeleteHandler;

            detached.CloseEvent += closeEventHandler;
            detached.CloseEvent -= closeEventHandler;

            detached.Open += openEventHandler;
            detached.Open -= openEventHandler;

            detached.PropertyChange += propertyChangeHandler;
            detached.PropertyChange -= propertyChangeHandler;

            detached.Read += readHandler;
            detached.Read -= readHandler;

            detached.Write += writeEventHandler;
            detached.Write -= writeEventHandler;
        }



        [TestMethod]
        public void Serialization_IgnoresComReferences()
        {
            var detached = new DetachedOutlookItem(CreateMockItem().Object);
            string json = JsonConvert.SerializeObject(detached);
            Assert.IsFalse(json.Contains("Application"));
            Assert.IsFalse(json.Contains("Attachments"));
            Assert.IsFalse(json.Contains("ItemProperties"));
            Assert.IsFalse(json.Contains("Session"));
            Assert.IsFalse(json.Contains("InnerObject"));
            Assert.IsFalse(json.Contains("\"Parent\""));
            Assert.IsTrue(json.Contains("BillingInformation")); // At least one value property
        }

        [TestMethod]
        public void Reattach_Throws_OnNullApplication()
        {
            var detached = new DetachedOutlookItem(CreateMockItem().Object);
            Assert.ThrowsException<ArgumentNullException>(() => detached.Reattach(null));
        }

        [TestMethod]
        public void Reattach_Throws_OnMissingEntryIDOrStoreID()
        {
            var mock = CreateMockItem();
            mock.SetupGet(m => m.EntryID).Returns((string)null);
            var detached = new DetachedOutlookItem(mock.Object);
            detached.StoreID = null;
            Assert.ThrowsException<InvalidOperationException>(() => detached.Reattach(new Mock<Application>().Object));
        }

        [TestMethod]
        public void Reattach_Throws_WhenStoreNotFound()
        {
            var detached = new DetachedOutlookItem(CreateMockItem().Object);
            var appMock = new Mock<Application>();
            var nsMock = new Mock<NameSpace>();
            appMock.Setup(a => a.Session).Returns(nsMock.Object);
            // Empty Stores collection
            var storesMock = new Mock<Stores>();
            storesMock.Setup(s => s.GetEnumerator()).Returns((System.Collections.IEnumerator)(new Store[0]).GetEnumerator());
            nsMock.Setup(n => n.Stores).Returns(storesMock.Object);
            Assert.ThrowsException<InvalidOperationException>(() => detached.Reattach(appMock.Object));
        }

        [TestMethod]
        public void Reattach_Throws_WhenItemNotFound()
        {
            var detached = new DetachedOutlookItem(CreateMockItem().Object);

            var appMock = new Mock<Application>();
            var nsMock = new Mock<NameSpace>();
            appMock.Setup(a => a.Session).Returns(nsMock.Object);

            var storeMock = new Mock<Store>();
            storeMock.SetupGet(s => s.StoreID).Returns(detached.StoreID);
            var storesMock = new Mock<Stores>();
            storesMock.Setup(s => s.GetEnumerator()).Returns((new[] { storeMock.Object }).GetEnumerator());
            nsMock.Setup(n => n.Stores).Returns(storesMock.Object);

            nsMock.Setup(n => n.GetItemFromID(detached.EntryID, detached.StoreID)).Returns((object)null);
            Assert.ThrowsException<InvalidOperationException>(() => detached.Reattach(appMock.Object));
        }

        [TestMethod]
        public void Reattach_ReturnsWrapper_WhenFound()
        {
            var detached = new DetachedOutlookItem(CreateMockItem().Object);

            var appMock = new Mock<Application>();
            var nsMock = new Mock<NameSpace>();
            appMock.Setup(a => a.Session).Returns(nsMock.Object);

            var storeMock = new Mock<Store>();
            storeMock.SetupGet(s => s.StoreID).Returns(detached.StoreID);
            var storesMock = new Mock<Stores>();
            storesMock.Setup(s => s.GetEnumerator()).Returns((new[] { storeMock.Object }).GetEnumerator());
            nsMock.Setup(n => n.Stores).Returns(storesMock.Object);

            var mailItemMock = new Mock<MailItem>().Object;
            nsMock.Setup(n => n.GetItemFromID(detached.EntryID, detached.StoreID)).Returns(mailItemMock);

            var wrapper = detached.Reattach(appMock.Object);
            Assert.IsNotNull(wrapper);
            Assert.AreSame(mailItemMock, wrapper.InnerObject);
        }
    }
}
