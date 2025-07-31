using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;

namespace Gsync.OutlookInterop.Interfaces.Items
{
    public interface IMailItem : IItem, IEquatable<IMailItem>
    {
        // --- Additional _MailItem properties ---
        string BCC { get; set; }
        string CC { get; set; }
        string DeferredDeliveryTime { get; set; }
        string DeleteAfterSubmit { get; set; }
        string FlagRequest { get; set; }
        string ReceivedByName { get; }
        string ReceivedOnBehalfOfName { get; }
        DateTime ReceivedTime { get; }
        string RecipientReassignmentProhibited { get; set; }
        Recipients Recipients { get; }
        bool ReminderOverrideDefault { get; set; }
        bool ReminderPlaySound { get; set; }
        bool ReminderSet { get; set; }
        string ReminderSoundFile { get; set; }
        DateTime ReminderTime { get; set; }
        string ReplyRecipientNames { get; }
        Recipients ReplyRecipients { get; }
        int SaveSentMessageFolder { get; set; }
        string SenderEmailType { get; }
        string SentOnBehalfOfName { get; set; }
        DateTime SentOn { get; }
        bool Submitted { get; }
        string To { get; set; }
        string VotingOptions { get; set; }
        string VotingResponse { get; set; }

        // --- Additional _MailItem methods ---
        void ClearConversationIndex();
        MailItem Forward();
        void ImportanceChanged();
        MailItem Reply();
        MailItem ReplyAll();
        void Send();

        // --- MailItem-specific events not in IItem ---
        //public virtual extern event ItemEvents_10_CustomActionEventHandler CustomAction;
        //public virtual extern event ItemEvents_10_CustomPropertyChangeEventHandler CustomPropertyChange;
        //public virtual extern event ItemEvents_10_ForwardEventHandler ItemEvents_10_Event_Forward;
        //public virtual extern event ItemEvents_10_ReplyEventHandler ItemEvents_10_Event_Reply;
        //public virtual extern event ItemEvents_10_ReplyAllEventHandler ItemEvents_10_Event_ReplyAll;
        //public virtual extern event ItemEvents_10_SendEventHandler ItemEvents_10_Event_Send;
        //public virtual extern event ItemEvents_10_BeforeCheckNamesEventHandler BeforeCheckNames;
        //public virtual extern event ItemEvents_10_BeforeAttachmentSaveEventHandler BeforeAttachmentSave;
        //public virtual extern event ItemEvents_10_BeforeAttachmentAddEventHandler BeforeAttachmentAdd;
        //public virtual extern event ItemEvents_10_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreview;
        //public virtual extern event ItemEvents_10_BeforeAttachmentReadEventHandler BeforeAttachmentRead;
        //public virtual extern event ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFile;
        //public virtual extern event ItemEvents_10_UnloadEventHandler Unload;
        //public virtual extern event ItemEvents_10_BeforeAutoSaveEventHandler BeforeAutoSave;
        //public virtual extern event ItemEvents_10_BeforeReadEventHandler BeforeRead;
        //public virtual extern event ItemEvents_10_AfterWriteEventHandler AfterWrite;
        //public virtual extern event ItemEvents_10_ReadCompleteEventHandler ReadComplete;

        event CustomActionEventHandler CustomAction;
        event CustomPropertyChangeEventHandler CustomPropertyChange;
        event ForwardEventHandler ForwardEvent;
        event ReplyEventHandler ReplyEvent;
        event ReplyAllEventHandler ReplyAllEvent;
        event SendEventHandler SendEvent;
        event BeforeCheckNamesEventHandler BeforeCheckNames;
        event BeforeAttachmentSaveEventHandler BeforeAttachmentSave;
        event BeforeAttachmentAddEventHandler BeforeAttachmentAdd;
        event UnloadEventHandler Unload;
        event BeforeAutoSaveEventHandler BeforeAutoSave;
        event BeforeReadEventHandler BeforeRead;

        // --- Event delegates for event bridging ---
        public delegate void CustomActionEventHandler(object Action, object Response, ref bool Cancel);
        public delegate void CustomPropertyChangeEventHandler(string Name);
        public delegate void ForwardEventHandler(object Forward, ref bool Cancel);
        public delegate void ReplyEventHandler(object Response, ref bool Cancel);
        public delegate void ReplyAllEventHandler(object Response, ref bool Cancel);
        public delegate void SendEventHandler(ref bool Cancel);
        public delegate void BeforeCheckNamesEventHandler(ref bool Cancel);
        public delegate void BeforeAttachmentSaveEventHandler(Attachment Attachment, ref bool Cancel);
        public delegate void BeforeAttachmentAddEventHandler(Attachment Attachment, ref bool Cancel);
        public delegate void UnloadEventHandler();
        public delegate void BeforeAutoSaveEventHandler(ref bool Cancel);
        public delegate void BeforeReadEventHandler();

        public new IEqualityComparer<IMailItem> EqualityComparer { get; set; }
    }
}
