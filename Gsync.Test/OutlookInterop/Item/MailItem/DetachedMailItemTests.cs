using Gsync.OutlookInterop.Interfaces.Items;
using Gsync.OutlookInterop.Item;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;

namespace Gsync.Test.OutlookInterop.Item
{
    [TestClass]
    public class DetachedMailItemTests
    {
        private Mock<IMailItem> CreateMockMailItem()
        {
            var mock = new Mock<IMailItem>();

            // Read-write properties (get; set;)
            mock.SetupProperty(x => x.BillingInformation, "BillInfo");
            mock.SetupProperty(x => x.Body, "Body");
            mock.SetupProperty(x => x.Categories, "Cat");
            mock.SetupProperty(x => x.Companies, "Co");
            mock.SetupProperty(x => x.Mileage, "100");
            mock.SetupProperty(x => x.NoAging, true);
            mock.SetupProperty(x => x.HTMLBody, "<html></html>");
            mock.SetupProperty(x => x.Importance, OlImportance.olImportanceHigh);
            mock.SetupProperty(x => x.ReminderOverrideDefault, true);
            mock.SetupProperty(x => x.ReminderPlaySound, false);
            mock.SetupProperty(x => x.ReminderSet, true);
            mock.SetupProperty(x => x.ReminderSoundFile, "sound.wav");
            mock.SetupProperty(x => x.ReminderTime, DateTime.Now.AddMinutes(30));
            mock.SetupProperty(x => x.SaveSentMessageFolder, 1);
            mock.SetupProperty(x => x.Sensitivity, OlSensitivity.olPrivate);
            mock.SetupProperty(x => x.Subject, "Subject");
            mock.SetupProperty(x => x.UnRead, false);

            // MailItem properties (get; set;)
            mock.SetupProperty(x => x.BCC, "bcc@example.com");
            mock.SetupProperty(x => x.CC, "cc@example.com");
            mock.SetupProperty(x => x.DeferredDeliveryTime, "2024-07-25T11:45:00Z");
            mock.SetupProperty(x => x.DeleteAfterSubmit, "Yes");
            mock.SetupProperty(x => x.FlagRequest, "Follow up");
            mock.SetupProperty(x => x.RecipientReassignmentProhibited, "No");
            mock.SetupProperty(x => x.SentOnBehalfOfName, "SentOnBehalf");
            mock.SetupProperty(x => x.To, "to@example.com");
            mock.SetupProperty(x => x.VotingOptions, "Yes;No");
            mock.SetupProperty(x => x.VotingResponse, "Yes");

            // Read-only properties (get;)
            mock.Setup(x => x.Class).Returns(OlObjectClass.olMail);
            mock.Setup(x => x.ConversationID).Returns("ConvId");
            mock.Setup(x => x.CreationTime).Returns(DateTime.Now.AddDays(-2));
            mock.Setup(x => x.EntryID).Returns("E123");
            mock.Setup(x => x.LastModificationTime).Returns(DateTime.Now);
            mock.Setup(x => x.MessageClass).Returns("IPM.Note");
            mock.Setup(x => x.OutlookInternalVersion).Returns(12345);
            mock.Setup(x => x.OutlookVersion).Returns("16.0");
            mock.Setup(x => x.Saved).Returns(true);
            mock.Setup(x => x.SenderEmailAddress).Returns("sender@example.com");
            mock.Setup(x => x.SenderName).Returns("Sender");
            mock.Setup(x => x.Size).Returns(1000);
            mock.Setup(x => x.ReceivedByName).Returns("ReceivedBy");
            mock.Setup(x => x.ReceivedOnBehalfOfName).Returns("ReceivedOnBehalf");
            mock.Setup(x => x.ReceivedTime).Returns(DateTime.Now.AddHours(-1));
            mock.Setup(x => x.ReplyRecipientNames).Returns("Rep1;Rep2");
            mock.Setup(x => x.SenderEmailType).Returns("SMTP");
            mock.Setup(x => x.SentOn).Returns(DateTime.Now.AddHours(-2));
            mock.Setup(x => x.Submitted).Returns(true);

            // Complex types: Recipients and ReplyRecipients (returning null is fine for detached tests)
            mock.Setup(x => x.Recipients).Returns((Recipients)null);
            mock.Setup(x => x.ReplyRecipients).Returns((Recipients)null);

            // Parent folder (for StoreID/ParentFolderEntryID testing)
            var folder = new Mock<MAPIFolder>();
            folder.Setup(f => f.StoreID).Returns("STORE_ID");
            folder.Setup(f => f.EntryID).Returns("FOLDER_ID");
            mock.Setup(x => x.Parent).Returns(folder.Object);

            return mock;
        }

        [TestMethod]
        public void Constructor_Parameterless_NotNull()
        {
            var detached = new DetachedMailItem();
            Assert.IsNotNull(detached);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_ThrowsOnNullArgument()
        {
            var detached = new DetachedMailItem(null);
        }

        [TestMethod]
        public void Properties_AreCopiedFromSource()
        {
            var mock = CreateMockMailItem();
            var item = mock.Object;
            var detached = new DetachedMailItem(item);

            // IItem (base) properties
            Assert.AreEqual(item.BillingInformation, detached.BillingInformation);
            Assert.AreEqual(item.Body, detached.Body);
            Assert.AreEqual(item.Categories, detached.Categories);
            Assert.AreEqual(item.Class, detached.Class);
            Assert.AreEqual(item.Companies, detached.Companies);
            Assert.AreEqual(item.ConversationID, detached.ConversationID);
            Assert.AreEqual(item.CreationTime, detached.CreationTime);
            Assert.AreEqual(item.EntryID, detached.EntryID);
            Assert.AreEqual(item.HTMLBody, detached.HTMLBody);
            Assert.AreEqual(item.Importance, detached.Importance);
            Assert.AreEqual(item.LastModificationTime, detached.LastModificationTime);
            Assert.AreEqual(item.MessageClass, detached.MessageClass);
            Assert.AreEqual(item.Mileage, detached.Mileage);
            Assert.AreEqual(item.NoAging, detached.NoAging);
            Assert.AreEqual(item.OutlookInternalVersion, detached.OutlookInternalVersion);
            Assert.AreEqual(item.OutlookVersion, detached.OutlookVersion);
            Assert.AreEqual(item.Saved, detached.Saved);
            Assert.AreEqual(item.SenderEmailAddress, detached.SenderEmailAddress);
            Assert.AreEqual(item.SenderName, detached.SenderName);
            Assert.AreEqual(item.Sensitivity, detached.Sensitivity);
            Assert.AreEqual(item.Size, detached.Size);
            Assert.AreEqual(item.Subject, detached.Subject);
            Assert.AreEqual(item.UnRead, detached.UnRead);
            Assert.AreEqual("STORE_ID", detached.StoreID);
            Assert.AreEqual("FOLDER_ID", detached.ParentFolderEntryID);

            // IMailItem properties
            Assert.AreEqual(item.BCC, detached.BCC);
            Assert.AreEqual(item.CC, detached.CC);
            Assert.AreEqual(item.DeferredDeliveryTime, detached.DeferredDeliveryTime);
            Assert.AreEqual(item.DeleteAfterSubmit, detached.DeleteAfterSubmit);
            Assert.AreEqual(item.FlagRequest, detached.FlagRequest);
            Assert.AreEqual(item.ReceivedByName, detached.ReceivedByName);
            Assert.AreEqual(item.ReceivedOnBehalfOfName, detached.ReceivedOnBehalfOfName);
            Assert.AreEqual(item.ReceivedTime, detached.ReceivedTime);
            Assert.AreEqual(item.RecipientReassignmentProhibited, detached.RecipientReassignmentProhibited);
            Assert.AreEqual(item.ReminderOverrideDefault, detached.ReminderOverrideDefault);
            Assert.AreEqual(item.ReminderPlaySound, detached.ReminderPlaySound);
            Assert.AreEqual(item.ReminderSet, detached.ReminderSet);
            Assert.AreEqual(item.ReminderSoundFile, detached.ReminderSoundFile);
            Assert.AreEqual(item.ReminderTime, detached.ReminderTime);
            Assert.AreEqual(item.ReplyRecipientNames, detached.ReplyRecipientNames);
            Assert.AreEqual(item.SaveSentMessageFolder, detached.SaveSentMessageFolder);
            Assert.AreEqual(item.SenderEmailType, detached.SenderEmailType);
            Assert.AreEqual(item.SentOnBehalfOfName, detached.SentOnBehalfOfName);
            Assert.AreEqual(item.SentOn, detached.SentOn);
            Assert.AreEqual(item.Submitted, detached.Submitted);
            Assert.AreEqual(item.To, detached.To);
            Assert.AreEqual(item.VotingOptions, detached.VotingOptions);
            Assert.AreEqual(item.VotingResponse, detached.VotingResponse);
        }

        [TestMethod]
        public void ComProperties_AreAlwaysNull()
        {
            var mock = CreateMockMailItem();
            var item = mock.Object;
            var detached = new DetachedMailItem(item);

            Assert.IsNull(detached.Application);
            Assert.IsNull(detached.Attachments);
            Assert.IsNull(detached.ItemProperties);
            Assert.IsNull(detached.Session);
            Assert.IsNull(detached.InnerObject);
            Assert.IsNull(detached.Parent);
            Assert.IsNull(detached.Recipients);
            Assert.IsNull(detached.ReplyRecipients);
        }

        [TestMethod]
        public void Methods_ThrowNotSupportedException()
        {
            var mock = CreateMockMailItem();
            var item = mock.Object;
            var detached = new DetachedMailItem(item);

            Assert.ThrowsException<NotSupportedException>(() => detached.Close(OlInspectorClose.olDiscard));
            Assert.ThrowsException<NotSupportedException>(() => detached.Copy());
            Assert.ThrowsException<NotSupportedException>(() => detached.Delete());
            Assert.ThrowsException<NotSupportedException>(() => detached.Display());
            Assert.ThrowsException<NotSupportedException>(() => detached.Move(null));
            Assert.ThrowsException<NotSupportedException>(() => detached.PrintOut());
            Assert.ThrowsException<NotSupportedException>(() => detached.Save());
            Assert.ThrowsException<NotSupportedException>(() => detached.SaveAs("x"));
            Assert.ThrowsException<NotSupportedException>(() => detached.ShowCategoriesDialog());

            Assert.ThrowsException<NotSupportedException>(() => detached.ClearConversationIndex());
            Assert.ThrowsException<NotSupportedException>(() => detached.Forward());
            Assert.ThrowsException<NotSupportedException>(() => detached.ImportanceChanged());
            Assert.ThrowsException<NotSupportedException>(() => detached.Reply());
            Assert.ThrowsException<NotSupportedException>(() => detached.ReplyAll());
            Assert.ThrowsException<NotSupportedException>(() => detached.Send());
        }

        [TestMethod]
        public void Events_SubscribeAndUnsubscribe_NoError()
        {
            var item = CreateMockMailItem().Object;
            var detached = new DetachedMailItem(item);

            // IItem events
            IItem.AttachmentAddEventHandler attachmentAdd = (Attachment a) => { };
            IItem.AttachmentReadEventHandler attachmentRead = (Attachment a) => { };
            IItem.AttachmentRemoveEventHandler attachmentRemove = (Attachment a) => { };
            IItem.BeforeDeleteEventHandler beforeDelete = (object obj, ref bool cancel) => { };
            IItem.CloseEventHandler close = (ref bool cancel) => { };
            IItem.OpenEventHandler open = (ref bool cancel) => { };
            IItem.PropertyChangeEventHandler propChange = (string name) => { };
            IItem.ReadEventHandler read = () => { };
            IItem.WriteEventHandler write = (ref bool cancel) => { };

            detached.AttachmentAdd += attachmentAdd; detached.AttachmentAdd -= attachmentAdd;
            detached.AttachmentRead += attachmentRead; detached.AttachmentRead -= attachmentRead;
            detached.AttachmentRemove += attachmentRemove; detached.AttachmentRemove -= attachmentRemove;
            detached.BeforeDelete += beforeDelete; detached.BeforeDelete -= beforeDelete;
            detached.CloseEvent += close; detached.CloseEvent -= close;
            detached.Open += open; detached.Open -= open;
            detached.PropertyChange += propChange; detached.PropertyChange -= propChange;
            detached.Read += read; detached.Read -= read;
            detached.Write += write; detached.Write -= write;

            // IMailItem events
            IMailItem.CustomActionEventHandler customAction = (object a, object r, ref bool c) => { };
            IMailItem.CustomPropertyChangeEventHandler customProp = (string s) => { };
            IMailItem.ForwardEventHandler forwardEv = (object forward, ref bool cancel) => { };
            IMailItem.ReplyEventHandler replyEv = (object response, ref bool cancel) => { };
            IMailItem.ReplyAllEventHandler replyAllEv = (object response, ref bool cancel) => { };
            IMailItem.SendEventHandler sendEv = (ref bool c) => { };
            IMailItem.BeforeCheckNamesEventHandler beforeCheckNames = (ref bool c) => { };
            IMailItem.BeforeAttachmentSaveEventHandler beforeAttachmentSave = (Attachment a, ref bool c) => { };
            IMailItem.BeforeAttachmentAddEventHandler beforeAttachmentAdd = (Attachment a, ref bool c) => { };
            IMailItem.UnloadEventHandler unload = () => { };
            IMailItem.BeforeAutoSaveEventHandler beforeAutoSave = (ref bool c) => { };
            IMailItem.BeforeReadEventHandler beforeRead = () => { };

            detached.CustomAction += customAction; detached.CustomAction -= customAction;
            detached.CustomPropertyChange += customProp; detached.CustomPropertyChange -= customProp;
            detached.ForwardEvent += forwardEv; detached.ForwardEvent -= forwardEv;
            detached.ReplyEvent += replyEv; detached.ReplyEvent -= replyEv;
            detached.ReplyAllEvent += replyAllEv; detached.ReplyAllEvent -= replyAllEv;
            detached.SendEvent += sendEv; detached.SendEvent -= sendEv;
            detached.BeforeCheckNames += beforeCheckNames; detached.BeforeCheckNames -= beforeCheckNames;
            detached.BeforeAttachmentSave += beforeAttachmentSave; detached.BeforeAttachmentSave -= beforeAttachmentSave;
            detached.BeforeAttachmentAdd += beforeAttachmentAdd; detached.BeforeAttachmentAdd -= beforeAttachmentAdd;
            detached.Unload += unload; detached.Unload -= unload;
            detached.BeforeAutoSave += beforeAutoSave; detached.BeforeAutoSave -= beforeAutoSave;
            detached.BeforeRead += beforeRead; detached.BeforeRead -= beforeRead;
        }

        // Helper to create a minimal valid DetachedMailItem
        private DetachedMailItem CreateDetached(string entryId = "ITEMID", string storeId = "STOREID")
        {
            return new DetachedMailItem
            {
                EntryID = entryId,
                StoreID = storeId
            };
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Reattach_ThrowsIfApplicationNull()
        {
            var detached = CreateDetached();
            detached.Reattach(null);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void Reattach_ThrowsIfEntryIDMissing()
        {
            var detached = CreateDetached(entryId: null, storeId: "STOREID");
            var appMock = new Mock<Application>();
            detached.Reattach(appMock.Object);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void Reattach_ThrowsIfStoreIDMissing()
        {
            var detached = CreateDetached(entryId: "ITEMID", storeId: null);
            var appMock = new Mock<Application>();
            detached.Reattach(appMock.Object);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void Reattach_ThrowsIfStoreNotFound()
        {
            // App.Session.Stores will not match STOREID
            var appMock = new Mock<Application>();
            var nsMock = new Mock<NameSpace>();
            var storesMock = new Mock<Stores>();
            storesMock.Setup(s => s.GetEnumerator()).Returns((System.Collections.IEnumerator)new Store[0].GetEnumerator());
            nsMock.SetupGet(ns => ns.Stores).Returns(storesMock.Object);
            appMock.SetupGet(a => a.Session).Returns(nsMock.Object);

            var detached = CreateDetached("ITEMID", "STOREID");
            detached.Reattach(appMock.Object);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void Reattach_ThrowsIfItemNotFound()
        {
            // App.Session.Stores has one matching store, but GetItemFromID returns null
            var appMock = new Mock<Application>();
            var nsMock = new Mock<NameSpace>();
            var storeMock = new Mock<Store>();
            storeMock.SetupGet(s => s.StoreID).Returns("STOREID");
            var storesList = new[] { storeMock.Object };
            var storesMock = new Mock<Stores>();
            storesMock.Setup(s => s.GetEnumerator()).Returns(storesList.GetEnumerator());
            nsMock.SetupGet(ns => ns.Stores).Returns(storesMock.Object);

            nsMock.Setup(ns => ns.GetItemFromID("ITEMID", "STOREID")).Returns((object)null);

            appMock.SetupGet(a => a.Session).Returns(nsMock.Object);

            var detached = CreateDetached("ITEMID", "STOREID");
            detached.Reattach(appMock.Object);
        }

        [TestMethod]
        public void Reattach_ReturnsMailItemWrapperIfFound()
        {
            // App.Session.Stores has one matching store, and GetItemFromID returns a mock MailItem
            var appMock = new Mock<Application>();
            var nsMock = new Mock<NameSpace>();
            var storeMock = new Mock<Store>();
            storeMock.SetupGet(s => s.StoreID).Returns("STOREID");
            var storesList = new[] { storeMock.Object };
            var storesMock = new Mock<Stores>();
            storesMock.Setup(s => s.GetEnumerator()).Returns(storesList.GetEnumerator());
            nsMock.SetupGet(ns => ns.Stores).Returns(storesMock.Object);

            var mailItem = new Mock<MailItem>().Object;
            nsMock.Setup(ns => ns.GetItemFromID("ITEMID", "STOREID")).Returns(mailItem);

            appMock.SetupGet(a => a.Session).Returns(nsMock.Object);

            var detached = CreateDetached("ITEMID", "STOREID");
            var result = detached.Reattach(appMock.Object);

            Assert.IsNotNull(result);
            Assert.IsInstanceOfType(result, typeof(MailItemWrapper)); // Or OutlookItemWrapper if that's the base
        }
    }
}
