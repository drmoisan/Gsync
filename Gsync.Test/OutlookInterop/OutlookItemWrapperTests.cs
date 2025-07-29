using Gsync.OutlookInterop.Interfaces.Items;
using Gsync.OutlookInterop.Item;
using Gsync.Utilities.HelperClasses;
using log4net.Appender;
using log4net.Core;
using log4net.Repository.Hierarchy;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Runtime.InteropServices;

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
        public void Close_WhenInvokeThrows_LogsError()
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
            //var wrapper = new OutlookItemWrapper(dummy);

            // Set up log4net memory appender
            var memoryAppender = new MemoryAppender();
            log4net.Config.BasicConfigurator.Configure(memoryAppender);

            // Act
            wrapper.Close(OlInspectorClose.olDiscard);

            // Assert: an error log was created
            var errorEvents = memoryAppender.GetEvents()
                .Where(ev => ev.Level == Level.Error &&
                             ev.RenderedMessage.Contains("Error closing item"))
                .ToList();

            Assert.IsTrue(errorEvents.Any(),
                "Expected error log for exception in Close.");

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

        [TestMethod]
        public void TryGet_LogsExceptionAndReturnsDefault()
        {
            var mock = CreateMailItemMock();
            // Simulate exception on property access            
            var mockException = new Mock<COMException>();
            mockException.SetupAllProperties();
            mockException.Setup(m => m.Message).Returns("Test exception").Verifiable();
            mockException.Setup(m => m.ErrorCode).Returns(80004005); // E_FAIL
            mock.Setup(m => m.Application).Throws(mockException.Object);            
            var wrapper = CreateWrapper(mock);

            // Should not throw, should return default (null for Application)
            var result = wrapper.Application;
            Assert.IsNull(result);
            mockException.Verify(m => m.Message, Times.AtLeastOnce);
        }

        [TestMethod]
        public void TrySet_LogsExceptionAndDoesNotThrow()
        {
            var mock = CreateMailItemMock();

            // Simulate exception on property set
            var mockException = new Mock<COMException>();
            mockException.SetupAllProperties();
            mockException.Setup(m => m.Message).Returns("Test exception").Verifiable();
            mockException.Setup(m => m.ErrorCode).Returns(80004005); // E_FAIL
            mock.SetupSet(m => m.Body = It.IsAny<string>()).Throws(mockException.Object);
            var wrapper = CreateWrapper(mock);

            // Should not throw
            wrapper.Body = "ShouldNotThrow";
            // No assertion needed: test passes if no exception is thrown
            mockException.Verify(m => m.Message, Times.AtLeastOnce);
        }

        [TestMethod]
        public void IsComObjectFunc_ReturnsFalse_ForNull()
        {
            var mock = CreateMailItemMock();
            var wrapper = CreateWrapper(mock);
            Assert.IsFalse(wrapperProtectedIsComObjectFunc(wrapper, null));
        }

        [TestMethod]
        public void IsComObjectFunc_ReturnsFalse_ForManagedObject()
        {
            var mock = CreateMailItemMock();
            var wrapper = CreateWrapper(mock);
            var managedObj = new object();
            Assert.IsFalse(wrapperProtectedIsComObjectFunc(wrapper, managedObj));
        }

        [TestMethod]
        public void IsComObjectFunc_ReturnsTrue_ForComObject()
        {
            var mock = CreateMailItemMock();
            var wrapper = CreateWrapper(mock);

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
                    Assert.IsTrue(wrapperProtectedIsComObjectFunc(wrapper, comObj));
                }
                finally
                {
                    Marshal.ReleaseComObject(comObj);
                }
            }
        }

        // Helper to access protected method
        private bool wrapperProtectedIsComObjectFunc(OutlookItemWrapper wrapper, object obj)
        {
            var method = typeof(OutlookItemWrapper).GetMethod("IsComObjectFunc", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            return (bool)method.Invoke(wrapper, new[] { obj });
        }

        [TestMethod]
        public void OnAttachmentAdd_RaisesEvent()
        {
            var mock = CreateMailItemMock();
            var wrapper = new TestableOutlookItemWrapper(mock.Object);
            bool called = false;
            var attachment = new Mock<Attachment>().Object;
            wrapper.AttachmentAdd += a =>
            {
                called = true;
                Assert.AreEqual(attachment, a);
            };

            wrapper.InvokeOnAttachmentAdd(attachment);
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnAttachmentRead_RaisesEvent()
        {
            var mock = CreateMailItemMock();
            var wrapper = new TestableOutlookItemWrapper(mock.Object);
            bool called = false;
            var attachment = new Mock<Attachment>().Object;
            wrapper.AttachmentRead += a =>
            {
                called = true;
                Assert.AreEqual(attachment, a);
            };

            wrapper.InvokeOnAttachmentRead(attachment);
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnAttachmentRemove_RaisesEvent()
        {
            var mock = CreateMailItemMock();
            var wrapper = new TestableOutlookItemWrapper(mock.Object);
            bool called = false;
            var attachment = new Mock<Attachment>().Object;
            wrapper.AttachmentRemove += a =>
            {
                called = true;
                Assert.AreEqual(attachment, a);
            };

            wrapper.InvokeOnAttachmentRemove(attachment);
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnBeforeDelete_RaisesEvent()
        {
            var mock = CreateMailItemMock();
            var wrapper = new TestableOutlookItemWrapper(mock.Object);
            bool called = false;
            object item = new object();
            bool cancel = false;
            wrapper.BeforeDelete += (object i, ref bool c) =>
            {
                called = true;
                Assert.AreEqual(item, i);
                c = true;
            };

            bool cancelArg = false;
            wrapper.InvokeOnBeforeDelete(item, ref cancelArg);
            Assert.IsTrue(called);
            Assert.IsTrue(cancelArg);
        }

        [TestMethod]
        public void OnCloseEvent_RaisesEvent()
        {
            var mock = CreateMailItemMock();
            var wrapper = new TestableOutlookItemWrapper(mock.Object);
            bool called = false;
            wrapper.CloseEvent += (ref bool c) =>
            {
                called = true;
                c = true;
            };

            bool cancel = false;
            wrapper.InvokeOnCloseEvent(ref cancel);
            Assert.IsTrue(called);
            Assert.IsTrue(cancel);
        }

        [TestMethod]
        public void OnOpen_RaisesEvent()
        {
            var mock = CreateMailItemMock();
            var wrapper = new TestableOutlookItemWrapper(mock.Object);
            bool called = false;
            wrapper.Open += (ref bool c) =>
            {
                called = true;
                c = true;
            };

            bool cancel = false;
            wrapper.InvokeOnOpen(ref cancel);
            Assert.IsTrue(called);
            Assert.IsTrue(cancel);
        }

        [TestMethod]
        public void OnPropertyChange_RaisesEvent()
        {
            var mock = CreateMailItemMock();
            var wrapper = new TestableOutlookItemWrapper(mock.Object);
            bool called = false;
            string propertyName = "TestProp";
            wrapper.PropertyChange += name =>
            {
                called = true;
                Assert.AreEqual(propertyName, name);
            };

            wrapper.InvokeOnPropertyChange(propertyName);
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnRead_RaisesEvent()
        {
            var mock = CreateMailItemMock();
            var wrapper = new TestableOutlookItemWrapper(mock.Object);
            bool called = false;
            wrapper.Read += () => called = true;

            wrapper.InvokeOnRead();
            Assert.IsTrue(called);
        }

        [TestMethod]
        public void OnWrite_RaisesEvent()
        {
            var mock = CreateMailItemMock();
            var wrapper = new TestableOutlookItemWrapper(mock.Object);
            bool called = false;
            wrapper.Write += (ref bool c) =>
            {
                called = true;
                c = true;
            };

            bool cancel = false;
            wrapper.InvokeOnWrite(ref cancel);
            Assert.IsTrue(called);
            Assert.IsTrue(cancel);
        }
                
    }

}