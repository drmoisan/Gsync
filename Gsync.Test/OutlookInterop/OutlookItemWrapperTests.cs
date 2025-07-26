using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Microsoft.Office.Interop.Outlook;
using Gsync.OutlookInterop.Item;
using Gsync.Utilities.HelperClasses;

namespace Gsync.Test.OutlookInterop.Item
{
    [TestClass]
    public class OutlookItemWrapperTests
    {
        [TestInitialize]
        public void Setup()
        {
            Console.SetOut(new DebugTextWriter());
        }

        private Mock<MailItem> CreateMailItemMock()
        {
            var mock = new Mock<MailItem>();
            mock.SetupAllProperties();
            return mock;
        }

        private OutlookItemWrapper CreateWrapper(Mock<MailItem> mock)
        {
            return new OutlookItemWrapper(mock.Object);
        }

        [TestMethod]
        public void Application_Property_ReturnsValue()
        {
            var mockMail = CreateMailItemMock();
            var app = new Mock<Application>().Object;
            mockMail.Setup(m => m.Application).Returns(app);
            var wrapper = CreateWrapper(mockMail);
            var app2 = wrapper.Application;
            Assert.AreEqual(app, wrapper.Application);
        }

        [TestMethod]
        public void Attachments_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            var attachments = new Mock<Attachments>().Object;
            mock.Setup(m => m.Attachments).Returns(attachments);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(attachments, wrapper.Attachments);
        }

        [TestMethod]
        public void BillingInformation_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.BillingInformation, "Test");
            var wrapper = CreateWrapper(mock);
            wrapper.BillingInformation = "NewValue";
            Assert.AreEqual("NewValue", wrapper.BillingInformation);
        }

        [TestMethod]
        public void Body_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.Body, "TestBody");
            var wrapper = CreateWrapper(mock);
            wrapper.Body = "NewBody";
            Assert.AreEqual("NewBody", wrapper.Body);
        }

        [TestMethod]
        public void Categories_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.Categories, "TestCat");
            var wrapper = CreateWrapper(mock);
            wrapper.Categories = "NewCat";
            Assert.AreEqual("NewCat", wrapper.Categories);
        }

        [TestMethod]
        public void Class_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.Class).Returns(OlObjectClass.olMail);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(OlObjectClass.olMail, wrapper.Class);
        }

        [TestMethod]
        public void Companies_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.Companies, "TestCo");
            var wrapper = CreateWrapper(mock);
            wrapper.Companies = "NewCo";
            Assert.AreEqual("NewCo", wrapper.Companies);
        }

        [TestMethod]
        public void ConversationID_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.ConversationID).Returns("ConvId");
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual("ConvId", wrapper.ConversationID);
        }

        [TestMethod]
        public void CreationTime_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            var dt = DateTime.Now;
            mock.Setup(m => m.CreationTime).Returns(dt);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(dt, wrapper.CreationTime);
        }

        [TestMethod]
        public void EntryID_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.EntryID).Returns("EntryId");
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual("EntryId", wrapper.EntryID);
        }

        [TestMethod]
        public void HTMLBody_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.HTMLBody, "TestHtml");
            var wrapper = CreateWrapper(mock);
            wrapper.HTMLBody = "NewHtml";
            Assert.AreEqual("NewHtml", wrapper.HTMLBody);
        }

        [TestMethod]
        public void Importance_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.Importance, OlImportance.olImportanceNormal);
            var wrapper = CreateWrapper(mock);
            wrapper.Importance = OlImportance.olImportanceHigh;
            Assert.AreEqual(OlImportance.olImportanceHigh, wrapper.Importance);
        }

        [TestMethod]
        public void ItemProperties_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            var itemProps = new Mock<ItemProperties>().Object;
            mock.Setup(m => m.ItemProperties).Returns(itemProps);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(itemProps, wrapper.ItemProperties);
        }

        [TestMethod]
        public void LastModificationTime_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            var dt = DateTime.Now;
            mock.Setup(m => m.LastModificationTime).Returns(dt);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(dt, wrapper.LastModificationTime);
        }

        [TestMethod]
        public void MessageClass_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.MessageClass).Returns("IPM.Note");
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual("IPM.Note", wrapper.MessageClass);
        }

        [TestMethod]
        public void Mileage_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.Mileage, "TestMileage");
            var wrapper = CreateWrapper(mock);
            wrapper.Mileage = "NewMileage";
            Assert.AreEqual("NewMileage", wrapper.Mileage);
        }

        [TestMethod]
        public void NoAging_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.NoAging, false);
            var wrapper = CreateWrapper(mock);
            wrapper.NoAging = true;
            Assert.IsTrue(wrapper.NoAging);
        }

        [TestMethod]
        public void OutlookInternalVersion_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.OutlookInternalVersion).Returns(1234);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(1234, wrapper.OutlookInternalVersion);
        }

        [TestMethod]
        public void OutlookVersion_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.OutlookVersion).Returns("16.0");
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual("16.0", wrapper.OutlookVersion);
        }

        [TestMethod]
        public void Parent_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            var parent = new object();
            mock.Setup(m => m.Parent).Returns(parent);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(parent, wrapper.Parent);
        }

        [TestMethod]
        public void Saved_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.Saved).Returns(true);
            var wrapper = CreateWrapper(mock);
            Assert.IsTrue(wrapper.Saved);
        }

        [TestMethod]
        public void SenderEmailAddress_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.SenderEmailAddress).Returns("test@example.com");
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual("test@example.com", wrapper.SenderEmailAddress);
        }

        [TestMethod]
        public void SenderName_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.SenderName).Returns("Sender");
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual("Sender", wrapper.SenderName);
        }

        [TestMethod]
        public void Sensitivity_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.Sensitivity, OlSensitivity.olNormal);
            var wrapper = CreateWrapper(mock);
            wrapper.Sensitivity = OlSensitivity.olConfidential;
            Assert.AreEqual(OlSensitivity.olConfidential, wrapper.Sensitivity);
        }

        [TestMethod]
        public void Session_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            var session = new Mock<NameSpace>().Object;
            mock.Setup(m => m.Session).Returns(session);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(session, wrapper.Session);
        }

        [TestMethod]
        public void Size_Property_ReturnsValue()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.Size).Returns(42);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(42, wrapper.Size);
        }

        [TestMethod]
        public void Subject_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.Subject, "TestSubject");
            var wrapper = CreateWrapper(mock);
            wrapper.Subject = "NewSubject";
            Assert.AreEqual("NewSubject", wrapper.Subject);
        }

        [TestMethod]
        public void UnRead_Property_GetSet()
        {
            var mock = CreateMailItemMock();
            mock.SetupProperty(m => m.UnRead, false);
            var wrapper = CreateWrapper(mock);
            wrapper.UnRead = true;
            Assert.IsTrue(wrapper.UnRead);
        }

        [TestMethod]
        public void Close_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.Close(It.IsAny<OlInspectorClose>())).Verifiable();
            var wrapper = CreateWrapper(mock);
            wrapper.Close(OlInspectorClose.olSave);
            mock.Verify(m => m.Close(OlInspectorClose.olSave), Times.Once);
        }

        [TestMethod]
        public void Copy_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            var copyObj = new object();
            mock.Setup(m => m.Copy()).Returns(copyObj);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(copyObj, wrapper.Copy());
        }

        [TestMethod]
        public void Delete_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.Delete()).Verifiable();
            var wrapper = CreateWrapper(mock);
            wrapper.Delete();
            mock.Verify(m => m.Delete(), Times.Once);
        }

        [TestMethod]
        public void Display_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.Display(It.IsAny<object>())).Verifiable();
            var wrapper = CreateWrapper(mock);
            wrapper.Display(true);
            mock.Verify(m => m.Display(true), Times.Once);
        }

        [TestMethod]
        public void Move_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            var folder = new Mock<MAPIFolder>().Object;
            var movedObj = new object();
            mock.Setup(m => m.Move(folder)).Returns(movedObj);
            var wrapper = CreateWrapper(mock);
            Assert.AreEqual(movedObj, wrapper.Move(folder));
        }

        [TestMethod]
        public void PrintOut_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.PrintOut()).Verifiable();
            var wrapper = CreateWrapper(mock);
            wrapper.PrintOut();
            mock.Verify(m => m.PrintOut(), Times.Once);
        }

        [TestMethod]
        public void Save_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.Save()).Verifiable();
            var wrapper = CreateWrapper(mock);
            wrapper.Save();
            mock.Verify(m => m.Save(), Times.Once);
        }

        [TestMethod]
        public void SaveAs_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.SaveAs(It.IsAny<string>(), It.IsAny<object>())).Verifiable();
            var wrapper = CreateWrapper(mock);
            wrapper.SaveAs("path", "type");
            mock.Verify(m => m.SaveAs("path", "type"), Times.Once);
        }

        [TestMethod]
        public void ShowCategoriesDialog_Method_CallsUnderlying()
        {
            var mock = CreateMailItemMock();
            mock.Setup(m => m.ShowCategoriesDialog()).Verifiable();
            var wrapper = CreateWrapper(mock);
            wrapper.ShowCategoriesDialog();
            mock.Verify(m => m.ShowCategoriesDialog(), Times.Once);
        }
    }
}