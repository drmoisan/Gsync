using Gsync.OutlookInterop.Interfaces.Items;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Gsync.OutlookInterop.Item;
using Moq;
using System;

namespace Gsync.Test.OutlookInterop.Item
{
    [TestClass]
    public class MailItemEqualityComparerTests
    {
        private Mock<IMailItem> CreateDefaultMailItemMock()
        {
            var mock = new Mock<IMailItem>();

            // Set up a canonical set of values
            mock.SetupProperty(x => x.BillingInformation, "BI");
            mock.SetupProperty(x => x.Body, "Body");
            mock.SetupProperty(x => x.Categories, "Cat");
            mock.SetupProperty(x => x.Companies, "Co");
            mock.SetupProperty(x => x.Mileage, "10");
            mock.SetupProperty(x => x.NoAging, true);
            mock.SetupProperty(x => x.HTMLBody, "<b>html</b>");
            mock.SetupProperty(x => x.Importance, OlImportance.olImportanceHigh);
            mock.SetupProperty(x => x.ReminderOverrideDefault, true);
            mock.SetupProperty(x => x.ReminderPlaySound, false);
            mock.SetupProperty(x => x.ReminderSet, true);
            mock.SetupProperty(x => x.ReminderSoundFile, "snd.wav");
            mock.SetupProperty(x => x.ReminderTime, DateTime.Today.AddHours(12));
            mock.SetupProperty(x => x.SaveSentMessageFolder, 42);
            mock.SetupProperty(x => x.Sensitivity, OlSensitivity.olPrivate);
            mock.SetupProperty(x => x.Subject, "Subject");
            mock.SetupProperty(x => x.UnRead, false);

            // MailItem properties
            mock.SetupProperty(x => x.BCC, "bcc@example.com");
            mock.SetupProperty(x => x.CC, "cc@example.com");
            mock.SetupProperty(x => x.DeferredDeliveryTime, "2024-07-25T11:45:00Z");
            mock.SetupProperty(x => x.DeleteAfterSubmit, "Yes");
            mock.SetupProperty(x => x.FlagRequest, "Flag");
            mock.SetupProperty(x => x.RecipientReassignmentProhibited, "No");
            mock.SetupProperty(x => x.SentOnBehalfOfName, "SenderB");
            mock.SetupProperty(x => x.To, "to@example.com");
            mock.SetupProperty(x => x.VotingOptions, "Yes;No");
            mock.SetupProperty(x => x.VotingResponse, "Yes");            

            // Read-only, set up via Setup (not SetupProperty)
            mock.Setup(x => x.Class).Returns(OlObjectClass.olMail);
            mock.Setup(x => x.ConversationID).Returns("ConvID");
            mock.Setup(x => x.CreationTime).Returns(DateTime.Today.AddDays(-1));
            mock.Setup(x => x.EntryID).Returns("E123");
            mock.Setup(x => x.LastModificationTime).Returns(DateTime.Today);
            mock.Setup(x => x.MessageClass).Returns("IPM.Note");
            mock.Setup(x => x.OutlookInternalVersion).Returns(12345);
            mock.Setup(x => x.OutlookVersion).Returns("16.0");
            mock.Setup(x => x.Saved).Returns(true);
            mock.Setup(x => x.SenderEmailAddress).Returns("sender@example.com");
            mock.Setup(x => x.SenderName).Returns("Sender");
            mock.Setup(x => x.Size).Returns(1000);

            mock.Setup(x => x.ReceivedByName).Returns("ReceivedBy");
            mock.Setup(x => x.ReceivedOnBehalfOfName).Returns("ReceivedOnBehalf");
            mock.Setup(x => x.ReceivedTime).Returns(DateTime.Today.AddHours(10));
            mock.Setup(x => x.ReplyRecipientNames).Returns("Rep1;Rep2");
            mock.Setup(x => x.SenderEmailType).Returns("SMTP");
            mock.Setup(x => x.SentOn).Returns(DateTime.Today.AddHours(8));
            mock.Setup(x => x.Submitted).Returns(true);            

            return mock;
        }

        [TestMethod]
        public void Equals_ReturnsTrue_ForReferenceEquals()
        {
            var mock = CreateDefaultMailItemMock();
            var comparer = new MailItemEqualityComparer();
            Assert.IsTrue(comparer.Equals(mock.Object, mock.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_IfEitherNull()
        {
            var mock = CreateDefaultMailItemMock();
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock.Object, null));
            Assert.IsFalse(comparer.Equals(null, mock.Object));
            Assert.IsTrue(comparer.Equals(null, null)); // This is .NET default; you can change this if desired.
        }

        [TestMethod]
        public void Equals_ReturnsTrue_ForIdenticalProperties()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            var comparer = new MailItemEqualityComparer();
            Assert.IsTrue(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_WhenAnyPropertyDiffers()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.Subject).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_IgnoresCase_ForStrings()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.Subject).Returns(mock1.Object.Subject.ToUpperInvariant());
            var comparer = new MailItemEqualityComparer();
            Assert.IsTrue(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void GetHashCode_IsEqual_ForEquivalentObjects()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            var comparer = new MailItemEqualityComparer();
            Assert.AreEqual(comparer.GetHashCode(mock1.Object), comparer.GetHashCode(mock2.Object));
        }

        [TestMethod]
        public void GetHashCode_Differs_WhenPropertyDiffers()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.Subject).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.AreNotEqual(comparer.GetHashCode(mock1.Object), comparer.GetHashCode(mock2.Object));
        }

        [TestMethod]
        public void GetHashCode_HandlesNulls()
        {
            var comparer = new MailItemEqualityComparer();
            Assert.AreEqual(0, comparer.GetHashCode(null));
        }

        [TestMethod]
        public void Equals_ReturnsTrue_WhenAllStringPropertiesAreNull()
        {
            var mock1 = new Mock<IMailItem>();
            var mock2 = new Mock<IMailItem>();
            // Set all properties to their default
            var comparer = new MailItemEqualityComparer();
            Assert.IsTrue(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_BillingInformation_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.BillingInformation, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Body_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.Body, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Categories_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.Categories, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Companies_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.Companies, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Mileage_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.Mileage, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_NoAging_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.NoAging, false);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_HTMLBody_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.HTMLBody, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Importance_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.Importance, OlImportance.olImportanceLow);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ReminderOverrideDefault_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.ReminderOverrideDefault, false);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ReminderPlaySound_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.ReminderPlaySound, true);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ReminderSet_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.ReminderSet, false);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ReminderSoundFile_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.ReminderSoundFile, "OTHER.WAV");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ReminderTime_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.ReminderTime, DateTime.Today.AddHours(1));
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_SaveSentMessageFolder_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.SaveSentMessageFolder, 999);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Sensitivity_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.Sensitivity, OlSensitivity.olConfidential);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Subject_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.Subject, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_UnRead_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.UnRead, true);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        // --- MailItem properties ---
        [TestMethod]
        public void Equals_ReturnsFalse_When_BCC_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.BCC, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_CC_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.CC, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_DeferredDeliveryTime_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.DeferredDeliveryTime, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_DeleteAfterSubmit_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.DeleteAfterSubmit, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_FlagRequest_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.FlagRequest, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_RecipientReassignmentProhibited_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.RecipientReassignmentProhibited, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_SentOnBehalfOfName_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.SentOnBehalfOfName, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_To_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.To, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_VotingOptions_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.VotingOptions, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_VotingResponse_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.SetupProperty(x => x.VotingResponse, "DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ConversationID_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.ConversationID).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_CreationTime_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.CreationTime).Returns(DateTime.Today.AddDays(-3));
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_EntryID_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.EntryID).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_LastModificationTime_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.LastModificationTime).Returns(DateTime.Today.AddDays(-1));
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_MessageClass_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.MessageClass).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_OutlookInternalVersion_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.OutlookInternalVersion).Returns(77777);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_OutlookVersion_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.OutlookVersion).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Saved_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.Saved).Returns(false);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_SenderEmailAddress_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.SenderEmailAddress).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_SenderName_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.SenderName).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Size_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.Size).Returns(9999);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ReceivedByName_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.ReceivedByName).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ReceivedOnBehalfOfName_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.ReceivedOnBehalfOfName).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ReceivedTime_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.ReceivedTime).Returns(DateTime.Today.AddHours(1));
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ReplyRecipientNames_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.ReplyRecipientNames).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_SenderEmailType_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.SenderEmailType).Returns("DIFFERENT");
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_SentOn_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.SentOn).Returns(DateTime.Today.AddHours(20));
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Submitted_Differs()
        {
            var mock1 = CreateDefaultMailItemMock();
            var mock2 = CreateDefaultMailItemMock();
            mock2.Setup(x => x.Submitted).Returns(false);
            var comparer = new MailItemEqualityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }
        
    }
}
