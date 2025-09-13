using Gsync.OutlookInterop.Interfaces.Items;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace Gsync.OutlookInterop.Item
{
    public class DetachedMailItem : DetachedOutlookItem, IMailItem
    {
        #region ctor

        public DetachedMailItem() : base() { }

        public DetachedMailItem(IMailItem item) : base(item)
        {
            if (item == null) throw new ArgumentNullException(nameof(item));

            BCC = item.BCC;
            CC = item.CC;
            DeferredDeliveryTime = item.DeferredDeliveryTime;
            DeleteAfterSubmit = item.DeleteAfterSubmit;
            FlagRequest = item.FlagRequest;
            HTMLBody = item.HTMLBody;
            ReceivedByName = item.ReceivedByName;
            ReceivedOnBehalfOfName = item.ReceivedOnBehalfOfName;
            ReceivedTime = item.ReceivedTime;
            RecipientReassignmentProhibited = item.RecipientReassignmentProhibited;
            ReminderOverrideDefault = item.ReminderOverrideDefault;
            ReminderPlaySound = item.ReminderPlaySound;
            ReminderSet = item.ReminderSet;
            ReminderSoundFile = item.ReminderSoundFile;
            ReminderTime = item.ReminderTime;
            ReplyRecipientNames = item.ReplyRecipientNames;
            SaveSentMessageFolder = item.SaveSentMessageFolder;
            SenderEmailAddress = item.SenderEmailAddress;
            SenderEmailType = item.SenderEmailType;
            SenderName = item.SenderName;
            SentOnBehalfOfName = item.SentOnBehalfOfName;
            SentOn = item.SentOn;
            Submitted = item.Submitted;
            To = item.To;
            VotingOptions = item.VotingOptions;
            VotingResponse = item.VotingResponse;
        }

        #endregion ctor

        #region DetachedMailItem Methods

        /// <summary>
        /// Reattach to a live MailItem if possible.
        /// </summary>
        public new MailItemWrapper Reattach(Application application)
        {
            if (application == null) throw new ArgumentNullException(nameof(application));
            if (string.IsNullOrWhiteSpace(this.EntryID) || string.IsNullOrWhiteSpace(this.StoreID))
                throw new InvalidOperationException("Insufficient information to reattach.");

            NameSpace session = application.Session;
            Store store = null;
            foreach (Store s in session.Stores)
            {
                if (s.StoreID == this.StoreID)
                {
                    store = s;
                    break;
                }
            }
            if (store == null)
                throw new InvalidOperationException("Store not found in this Outlook session.");

            object comObject = session.GetItemFromID(this.EntryID, this.StoreID);
            if (comObject == null)
                throw new InvalidOperationException("Outlook item not found in current store/session.");

            return new MailItemWrapper(comObject);
        }        
        
        #endregion DetachedMailItem Methods

        #region IMailItem Implementation

        #region IMailItem Properties

        // --- IMailItem value/scalar properties ---
        public string BCC { get; set; }
        public string CC { get; set; }
        public string DeferredDeliveryTime { get; set; }
        public string DeleteAfterSubmit { get; set; }
        public string FlagRequest { get; set; }
        public string HTMLBody { get; set; }
        public string ReceivedByName { get; set; }
        public string ReceivedOnBehalfOfName { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string RecipientReassignmentProhibited { get; set; }
        public bool ReminderOverrideDefault { get; set; }
        public bool ReminderPlaySound { get; set; }
        public bool ReminderSet { get; set; }
        public string ReminderSoundFile { get; set; }
        public DateTime ReminderTime { get; set; }
        public string ReplyRecipientNames { get; set; }
        public int SaveSentMessageFolder { get; set; }
        public string SenderEmailAddress { get; set; }
        public string SenderEmailType { get; set; }
        public string SenderName { get; set; }
        public string SentOnBehalfOfName { get; set; }
        public DateTime SentOn { get; set; }
        public bool Submitted { get; set; }
        public string To { get; set; }
        public string VotingOptions { get; set; }
        public string VotingResponse { get; set; }

        [JsonIgnore] public Recipients Recipients => null;
        [JsonIgnore] public Recipients ReplyRecipients => null;
        
        #endregion IMailItem Properties

        #region IMailItem Methods

        // --- IMailItem methods: NotSupportedException ---
        public void ClearConversationIndex() =>
            throw new NotSupportedException("This DetachedMailItem is not connected to Outlook.");
        public MailItem Forward() =>
            throw new NotSupportedException("This DetachedMailItem is not connected to Outlook.");
        public void ImportanceChanged() =>
            throw new NotSupportedException("This DetachedMailItem is not connected to Outlook.");
        public MailItem Reply() =>
            throw new NotSupportedException("This DetachedMailItem is not connected to Outlook.");
        public MailItem ReplyAll() =>
            throw new NotSupportedException("This DetachedMailItem is not connected to Outlook.");
        public void Send() =>
            throw new NotSupportedException("This DetachedMailItem is not connected to Outlook.");
        public void ShowCategoriesDialog()
        {
            throw new NotSupportedException("This DetachedMailItem is not connected to Outlook.");
        }
        
        #endregion IMailItem Methods

        #region IMailItem Events

        // --- IMailItem events: No-op ---
        public event IMailItem.CustomActionEventHandler CustomAction { add { } remove { } }
        public event IMailItem.CustomPropertyChangeEventHandler CustomPropertyChange { add { } remove { } }
        public event IMailItem.ForwardEventHandler ForwardEvent { add { } remove { } }
        public event IMailItem.ReplyEventHandler ReplyEvent { add { } remove { } }
        public event IMailItem.ReplyAllEventHandler ReplyAllEvent { add { } remove { } }
        public event IMailItem.SendEventHandler SendEvent { add { } remove { } }
        public event IMailItem.BeforeCheckNamesEventHandler BeforeCheckNames { add { } remove { } }
        public event IMailItem.BeforeAttachmentSaveEventHandler BeforeAttachmentSave { add { } remove { } }
        public event IMailItem.BeforeAttachmentAddEventHandler BeforeAttachmentAdd { add { } remove { } }
        public event IMailItem.UnloadEventHandler Unload { add { } remove { } }
        public event IMailItem.BeforeAutoSaveEventHandler BeforeAutoSave { add { } remove { } }
        public event IMailItem.BeforeReadEventHandler BeforeRead { add { } remove { } }
        
        #endregion IMailItem Events

        #endregion IMailItem Implementation

        #region IEquatable<IMailItem> Implementation

        private IEqualityComparer<IMailItem> _equalityComparer = new IItemEqualityComparer();
        /// <summary>
        /// Gets or sets the equality comparer used for IEquatable<IItem> implementation.
        /// </summary>
        public new IEqualityComparer<IMailItem> EqualityComparer
        {
            get => _equalityComparer;
            set => _equalityComparer = value ?? new IItemEqualityComparer();
        }

#nullable enable

        /// <summary>
        /// Implements IEquatable<IItem> using the injected or default IEqualityComparer<IItem>.
        /// </summary>
        public bool Equals(IMailItem? other)
        {
            return EqualityComparer.Equals(this, other);
        }

        public override bool Equals(object? obj)
        {
            if (ReferenceEquals(this, obj)) return true;
            if (obj is IMailItem item)
                return Equals(item);
            return false;
        }

        public override int GetHashCode()
        {
            return EqualityComparer.GetHashCode(this);
        }

#nullable disable

        #endregion IEquatable<IMailItem> Implementation
    }
}
