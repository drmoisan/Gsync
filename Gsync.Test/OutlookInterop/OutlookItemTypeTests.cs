using Gsync.OutlookInterop.Item;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Microsoft.Office.Interop.Outlook;

namespace Gsync.Tests.OutlookInterop.Item
{
    [TestClass]
    public class OutlookItemTypeTests
    {
        [TestMethod]
        public void GetType_ReturnsNull_WhenObjectIsNull()
        {
            Assert.IsNull(OutlookItemType.GetType(null));
        }

        [TestMethod]
        public void GetType_ReturnsNull_WhenObjectIsUnknownType()
        {
            var mock = new Mock<object>();
            //Assert.IsNull(OutlookItemType.GetType(mock.Object));
            Assert.AreEqual(mock.Object.GetType(), OutlookItemType.GetType(mock.Object));
        }

        [TestMethod]
        public void GetType_ReturnsAppointmentItem() =>
            Assert.AreEqual(typeof(AppointmentItem).FullName, OutlookItemType.GetType(new Mock<AppointmentItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsContactItem() =>
            Assert.AreEqual(typeof(ContactItem).FullName, OutlookItemType.GetType(new Mock<ContactItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsDistListItem() =>
            Assert.AreEqual(typeof(DistListItem).FullName, OutlookItemType.GetType(new Mock<DistListItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsDocumentItem() =>
            Assert.AreEqual(typeof(DocumentItem).FullName, OutlookItemType.GetType(new Mock<DocumentItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsJournalItem() =>
            Assert.AreEqual(typeof(JournalItem).FullName, OutlookItemType.GetType(new Mock<JournalItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsMailItem() =>
            Assert.AreEqual(typeof(MailItem).FullName, OutlookItemType.GetType(new Mock<MailItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsMeetingItem() =>
            Assert.AreEqual(typeof(MeetingItem).FullName, OutlookItemType.GetType(new Mock<MeetingItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsNoteItem() =>
            Assert.AreEqual(typeof(NoteItem).FullName, OutlookItemType.GetType(new Mock<NoteItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsPostItem() =>
            Assert.AreEqual(typeof(PostItem).FullName, OutlookItemType.GetType(new Mock<PostItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsRemoteItem() =>
            Assert.AreEqual(typeof(RemoteItem).FullName, OutlookItemType.GetType(new Mock<RemoteItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsReportItem() =>
            Assert.AreEqual(typeof(ReportItem).FullName, OutlookItemType.GetType(new Mock<ReportItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsSharingItem() =>
            Assert.AreEqual(typeof(SharingItem).FullName, OutlookItemType.GetType(new Mock<SharingItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsStorageItem() =>
            Assert.AreEqual(typeof(StorageItem).FullName, OutlookItemType.GetType(new Mock<StorageItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsTaskItem() =>
            Assert.AreEqual(typeof(TaskItem).FullName, OutlookItemType.GetType(new Mock<TaskItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsTaskRequestAcceptItem() =>
            Assert.AreEqual(typeof(TaskRequestAcceptItem).FullName, OutlookItemType.GetType(new Mock<TaskRequestAcceptItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsTaskRequestDeclineItem() =>
            Assert.AreEqual(typeof(TaskRequestDeclineItem).FullName, OutlookItemType.GetType(new Mock<TaskRequestDeclineItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsTaskRequestItem() =>
            Assert.AreEqual(typeof(TaskRequestItem).FullName, OutlookItemType.GetType(new Mock<TaskRequestItem>().Object).FullName);
        [TestMethod]
        public void GetType_ReturnsTaskRequestUpdateItem() =>
            Assert.AreEqual(typeof(TaskRequestUpdateItem).FullName, OutlookItemType.GetType(new Mock<TaskRequestUpdateItem>().Object).FullName);
    }
}
