using Gsync.OutlookInterop.Interfaces.Items;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Gsync.OutlookInterop.Item;
using Moq;
using System;

namespace Gsync.Test.OutlookInterop.Item
{
    [TestClass]
    public class IItemSimilarityComparerTests
    {
        private Mock<IItem> CreateDefaultItemMock()
        {
            var mock = new Mock<IItem>();

            // Set up a canonical set of values
            mock.SetupProperty(x => x.BillingInformation, "BI");
            mock.SetupProperty(x => x.Body, "Body");
            mock.SetupProperty(x => x.Categories, "Cat");
            mock.SetupProperty(x => x.Companies, "Co");
            mock.SetupProperty(x => x.Mileage, "10");
            mock.SetupProperty(x => x.NoAging, true);
            mock.SetupProperty(x => x.HTMLBody, "<b>html</b>");
            mock.SetupProperty(x => x.Importance, OlImportance.olImportanceHigh);
            mock.SetupProperty(x => x.Sensitivity, OlSensitivity.olPrivate);
            mock.SetupProperty(x => x.Subject, "Subject");
            mock.SetupProperty(x => x.UnRead, false);

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

            return mock;
        }

        [TestMethod]
        public void Equals_ReturnsTrue_ForReferenceEquals()
        {
            var mock = CreateDefaultItemMock();
            var comparer = new IItemSimilarityComparer();
            Assert.IsTrue(comparer.Equals(mock.Object, mock.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_IfEitherNull()
        {
            var mock = CreateDefaultItemMock();
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock.Object, null));
            Assert.IsFalse(comparer.Equals(null, mock.Object));
            Assert.IsTrue(comparer.Equals(null, null));
        }

        [TestMethod]
        public void Equals_ReturnsTrue_ForIdenticalProperties()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            var comparer = new IItemSimilarityComparer();
            Assert.IsTrue(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_WhenAnyPropertyDiffers_ExceptEntryID()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Subject, "DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_IgnoresCase_ForStrings()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Subject, mock1.Object.Subject.ToUpperInvariant());
            var comparer = new IItemSimilarityComparer();
            Assert.IsTrue(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void GetHashCode_IsEqual_ForEquivalentObjects()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            var comparer = new IItemSimilarityComparer();
            Assert.AreEqual(comparer.GetHashCode(mock1.Object), comparer.GetHashCode(mock2.Object));
        }

        [TestMethod]
        public void GetHashCode_Differs_WhenPropertyDiffers()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Subject, "DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.AreNotEqual(comparer.GetHashCode(mock1.Object), comparer.GetHashCode(mock2.Object));
        }

        [TestMethod]
        public void GetHashCode_HandlesNulls()
        {
            var comparer = new IItemSimilarityComparer();
            Assert.AreEqual(0, comparer.GetHashCode(null));
        }

        [TestMethod]
        public void Equals_ReturnsTrue_WhenAllStringPropertiesAreNull()
        {
            var mock1 = new Mock<IItem>();
            var mock2 = new Mock<IItem>();
            var comparer = new IItemSimilarityComparer();
            Assert.IsTrue(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_BillingInformation_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.BillingInformation, "DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Body_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Body, "DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Categories_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Categories, "DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Companies_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Companies, "DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Mileage_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Mileage, "DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_NoAging_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.NoAging, false);
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_HTMLBody_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.HTMLBody, "DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Importance_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Importance, OlImportance.olImportanceLow);
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Sensitivity_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Sensitivity, OlSensitivity.olConfidential);
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Subject_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.Subject, "DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_UnRead_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.SetupProperty(x => x.UnRead, true);
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Class_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.Class).Returns(OlObjectClass.olAppointment);
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_ConversationID_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.ConversationID).Returns("DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_CreationTime_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.CreationTime).Returns(DateTime.Today.AddDays(-3));
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsTrue_When_EntryID_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.EntryID).Returns("DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            // EntryID is ignored in similarity
            Assert.IsTrue(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_MessageClass_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.MessageClass).Returns("DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_OutlookInternalVersion_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.OutlookInternalVersion).Returns(77777);
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_OutlookVersion_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.OutlookVersion).Returns("DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Saved_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.Saved).Returns(false);
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_SenderEmailAddress_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.SenderEmailAddress).Returns("DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_SenderName_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.SenderName).Returns("DIFFERENT");
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }

        [TestMethod]
        public void Equals_ReturnsFalse_When_Size_Differs()
        {
            var mock1 = CreateDefaultItemMock();
            var mock2 = CreateDefaultItemMock();
            mock2.Setup(x => x.Size).Returns(9999);
            var comparer = new IItemSimilarityComparer();
            Assert.IsFalse(comparer.Equals(mock1.Object, mock2.Object));
        }
    }
}