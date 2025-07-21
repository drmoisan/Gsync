using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Net.Mail;
using Gsync;
using System.Linq;

namespace Gsync.Test
{
    [TestClass]
    public class ImapOutlookItemWrapperTests
    {
        [TestMethod]
        public void Equals_WithIdenticalMessageId_ReturnsTrue()
        {
            var item1 = new ImapOutlookItemWrapper(
                "msgid-123", "Subject", "from@example.com", "to@example.com", new DateTimeOffset(2023, 1, 1, 0, 0, 0, TimeSpan.Zero), "1");
            var item2 = new ImapOutlookItemWrapper(
                "msgid-123", "Different Subject", "otherfrom@example.com", "otherto@example.com", new DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero), "2");

            Assert.IsTrue(item1.Equals(item2));
        }

        [TestMethod]
        public void Equals_WithDifferentMessageIdAndSameOtherProperties_ReturnsFalse()
        {
            var date = new DateTimeOffset(2023, 1, 1, 0, 0, 0, TimeSpan.Zero);
            var item1 = new ImapOutlookItemWrapper(
                "msgid-1", "Subject", "from@example.com", "to@example.com", date, "1");
            var item2 = new ImapOutlookItemWrapper(
                "msgid-2", "Subject", "from@example.com", "to@example.com", date, "2");

            Assert.IsFalse(item1.Equals(item2));
        }

        [TestMethod]
        public void Equals_WithNullMessageIdAndSameOtherProperties_ReturnsTrue()
        {
            var date = new DateTimeOffset(2023, 1, 1, 0, 0, 0, TimeSpan.Zero);
            var item1 = new ImapOutlookItemWrapper(
                null, "Subject", "from@example.com", "to@example.com", date, "1");
            var item2 = new ImapOutlookItemWrapper(
                null, "Subject", "from@example.com", "to@example.com", date, "2");

            Assert.IsTrue(item1.Equals(item2));
        }

        [TestMethod]
        public void Equals_WithNullMessageIdAndSameOtherPropertiesAsObject_ReturnsTrue()
        {
            var date = new DateTimeOffset(2023, 1, 1, 0, 0, 0, TimeSpan.Zero);
            var item1 = new ImapOutlookItemWrapper(
                null, "Subject", "from@example.com", "to@example.com", date, "1");
            var item2 = new ImapOutlookItemWrapper(
                null, "Subject", "from@example.com", "to@example.com", date, "2");

            Assert.IsTrue(item1.Equals(item2 as object));
        }

        [TestMethod]
        public void Constructor_FromMailMessage_ExtractsPropertiesCorrectly()
        {
            var mailMessage = new MailMessage();
            mailMessage.Subject = "Test Subject";
            mailMessage.From = new MailAddress("from@example.com");
            mailMessage.To.Add(new MailAddress("to@example.com"));
            mailMessage.Headers.Add("Message-ID", "<msgid-xyz>");
            mailMessage.Headers.Add("Date", "Mon, 01 Jan 2024 12:00:00 +0000");

            var wrapper = new ImapOutlookItemWrapper(mailMessage, "imap-uid-123");

            Assert.AreEqual("<msgid-xyz>", wrapper.MessageId);
            Assert.AreEqual("Test Subject", wrapper.Subject);
            Assert.AreEqual("from@example.com", wrapper.From);
            Assert.AreEqual("to@example.com", wrapper.To);
            Assert.AreEqual(new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.Zero), wrapper.Date);
            Assert.AreEqual("imap-uid-123", wrapper.ImapUid);
        }

        [TestMethod]
        public void Properties_AreSetCorrectly()
        {
            var date = new DateTimeOffset(2022, 5, 10, 15, 30, 0, TimeSpan.FromHours(-4));
            var wrapper = new ImapOutlookItemWrapper("msgid-abc", "A subject", "a@b.com", "b@c.com", date, "uid-42");

            Assert.AreEqual("msgid-abc", wrapper.MessageId);
            Assert.AreEqual("A subject", wrapper.Subject);
            Assert.AreEqual("a@b.com", wrapper.From);
            Assert.AreEqual("b@c.com", wrapper.To);
            Assert.AreEqual(date, wrapper.Date);
            Assert.AreEqual("uid-42", wrapper.ImapUid);
        }

        [TestMethod]
        public void Equals_ObjectOverride_Works()
        {
            var date = new DateTimeOffset(2023, 1, 1, 0, 0, 0, TimeSpan.Zero);
            var item1 = new ImapOutlookItemWrapper("msgid-1", "Subject", "from@example.com", "to@example.com", date, "1");
            object item2 = new ImapOutlookItemWrapper("msgid-1", "Subject", "from@example.com", "to@example.com", date, "1");

            Assert.IsTrue(item1.Equals(item2));
        }

        [TestMethod]
        public void GetHashCode_UsesMessageId_WhenAvailable()
        {
            var item1 = new ImapOutlookItemWrapper("msgid-xyz", "Subject", "from@example.com", "to@example.com", new DateTimeOffset(2023, 1, 1, 0, 0, 0, TimeSpan.Zero), "1");
            var item2 = new ImapOutlookItemWrapper("msgid-xyz", "Other", "other@example.com", "other2@example.com", new DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero), "2");

            Assert.AreEqual(item1.GetHashCode(), item2.GetHashCode());
        }

        [TestMethod]
        public void GetHashCode_UsesOtherProperties_WhenMessageIdMissing()
        {
            var date = new DateTimeOffset(2023, 1, 1, 0, 0, 0, TimeSpan.Zero);
            var item1 = new ImapOutlookItemWrapper(null, "Subject", "from@example.com", "to@example.com", date, "1");
            var item2 = new ImapOutlookItemWrapper(null, "Subject", "from@example.com", "to@example.com", date, "2");

            Assert.AreEqual(item1.GetHashCode(), item2.GetHashCode());
        }

        [TestMethod]
        public void ParseDateHeader_InvalidDate_ReturnsMinValue()
        {
            // Use reflection to access the private static method
            var method = typeof(ImapOutlookItemWrapper).GetMethod("ParseDateHeader", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            var result = (DateTimeOffset)method.Invoke(null, new object[] { "not a date" });
            Assert.AreEqual(DateTimeOffset.MinValue, result);
        }

        [TestMethod]
        public void ParseDateHeader_ValidDateWithOffset_ReturnsCorrectValue()
        {
            var method = typeof(ImapOutlookItemWrapper).GetMethod("ParseDateHeader", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            var result = (DateTimeOffset)method.Invoke(null, new object[] { "Mon, 01 Jan 2024 12:00:00 +0200" });
            Assert.AreEqual(new DateTimeOffset(2024, 1, 1, 12, 0, 0, TimeSpan.FromHours(2)), result);
        }
    }
}
