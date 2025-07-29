using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Immutable;
using System.Reflection;
using Gsync.OutlookInterop.Interfaces.Items;
using Gsync.Utilities.Extensions;

namespace Gsync.OutlookInterop.Item
{
    public class MailItemWrapper : OutlookItemWrapper, IMailItem
    {
        private MailItem _mailItem;
        private bool _mailItemEventsAttached = false;

        // --- Constructors ---

        public MailItemWrapper(object item)
            : base(item)
        {
            _mailItem = (item as MailItem).ThrowIfNull($"Cannot cast item to {nameof(MailItem)}");
            AttachMailItemEvents();
        }

        protected MailItemWrapper(object item, ItemEvents_10_Event comEvents)
            : base(item, comEvents)
        {
            _mailItem = (item as MailItem).ThrowIfNull($"Cannot cast item to {nameof(MailItem)}");
            AttachMailItemEvents();
        }

        protected MailItemWrapper(object item, ItemEvents_10_Event comEvents, ImmutableHashSet<string> supportedTypes)
            : base(item, comEvents, supportedTypes)
        {
            _mailItem = (item as MailItem).ThrowIfNull($"Cannot cast item to {nameof(MailItem)}");
            AttachMailItemEvents();
        }

        public MailItemWrapper(OutlookItemWrapper baseWrapper)
            : base(baseWrapper.InnerObject, GetCom10Events(baseWrapper), baseWrapper.SupportedTypes)
        {
            _mailItem = (baseWrapper.InnerObject as MailItem).ThrowIfNull($"Cannot cast item to {nameof(MailItem)}");
            AttachMailItemEvents();
        }
                
        private static ItemEvents_10_Event GetCom10Events(OutlookItemWrapper wrapper)
        {
            var field = typeof(OutlookItemWrapper).GetField("_comEvents", BindingFlags.NonPublic | BindingFlags.Instance);
            return (ItemEvents_10_Event)field?.GetValue(wrapper);
        }

        // --- IMailItem Properties (fully forwarded, exception on COM error) ---

        public string BCC
        {
            get => _dyn.BCC;
            set => _dyn.BCC = value;
        }
        public string CC
        {
            get => _dyn.CC;
            set => _dyn.CC = value;
        }
        public string DeferredDeliveryTime
        {
            get => _dyn.DeferredDeliveryTime;
            set => _dyn.DeferredDeliveryTime = value;
        }
        public string DeleteAfterSubmit
        {
            get => _dyn.DeleteAfterSubmit;
            set => _dyn.DeleteAfterSubmit = value;
        }
        public string FlagRequest
        {
            get => _dyn.FlagRequest;
            set => _dyn.FlagRequest = value;
        }
        public string ReceivedByName => _dyn.ReceivedByName;
        public string ReceivedOnBehalfOfName => _dyn.ReceivedOnBehalfOfName;
        public DateTime ReceivedTime => _dyn.ReceivedTime;
        public string RecipientReassignmentProhibited
        {
            get => _dyn.RecipientReassignmentProhibited;
            set => _dyn.RecipientReassignmentProhibited = value;
        }
        public Recipients Recipients => _dyn.Recipients;
        public bool ReminderOverrideDefault
        {
            get => _dyn.ReminderOverrideDefault;
            set => _dyn.ReminderOverrideDefault = value;
        }
        public bool ReminderPlaySound
        {
            get => _dyn.ReminderPlaySound;
            set => _dyn.ReminderPlaySound = value;
        }
        public bool ReminderSet
        {
            get => _dyn.ReminderSet;
            set => _dyn.ReminderSet = value;
        }
        public string ReminderSoundFile
        {
            get => _dyn.ReminderSoundFile;
            set => _dyn.ReminderSoundFile = value;
        }
        public DateTime ReminderTime
        {
            get => _dyn.ReminderTime;
            set => _dyn.ReminderTime = value;
        }
        public string ReplyRecipientNames => _dyn.ReplyRecipientNames;
        public Recipients ReplyRecipients => _dyn.ReplyRecipients;
        public int SaveSentMessageFolder
        {
            get => _dyn.SaveSentMessageFolder;
            set => _dyn.SaveSentMessageFolder = value;
        }
        public string SenderEmailType => _dyn.SenderEmailType;
        public string SentOnBehalfOfName
        {
            get => _dyn.SentOnBehalfOfName;
            set => _dyn.SentOnBehalfOfName = value;
        }
        public DateTime SentOn => _dyn.SentOn;
        public bool Submitted => _dyn.Submitted;
        public string To
        {
            get => _dyn.To;
            set => _dyn.To = value;
        }
        public string VotingOptions
        {
            get => _dyn.VotingOptions;
            set => _dyn.VotingOptions = value;
        }
        public string VotingResponse
        {
            get => _dyn.VotingResponse;
            set => _dyn.VotingResponse = value;
        }

        // --- IMailItem Methods (fully forwarded, exception on COM error) ---

        public void ClearConversationIndex() => _dyn.ClearConversationIndex();
        public MailItem Forward() => _dyn.Forward();
        public void ImportanceChanged() => _dyn.ImportanceChanged();
        public MailItem Reply() => _dyn.Reply();
        public MailItem ReplyAll() => _dyn.ReplyAll();
        public void Send() => _dyn.Send();

        // --- IMailItem Events ---

        public event IMailItem.CustomActionEventHandler CustomAction;
        public event IMailItem.CustomPropertyChangeEventHandler CustomPropertyChange;
        public event IMailItem.ForwardEventHandler ForwardEvent;
        public event IMailItem.ReplyEventHandler ReplyEvent;
        public event IMailItem.ReplyAllEventHandler ReplyAllEvent;
        public event IMailItem.SendEventHandler SendEvent;
        public event IMailItem.BeforeCheckNamesEventHandler BeforeCheckNames;
        public event IMailItem.BeforeAttachmentSaveEventHandler BeforeAttachmentSave;
        public event IMailItem.BeforeAttachmentAddEventHandler BeforeAttachmentAdd;
        public event IMailItem.UnloadEventHandler Unload;
        public event IMailItem.BeforeAutoSaveEventHandler BeforeAutoSave;
        public event IMailItem.BeforeReadEventHandler BeforeRead;

        // --- Attach/Detach pattern for event bridging ---

        private void AttachMailItemEvents()
        {
            if (_mailItemEventsAttached) return;
            _mailItem = _item as MailItem;
            if (_mailItem == null) return;

            //var mie = (MailItemEvents_Event)_mailItem;
            var mie = base._comEvents;

            mie.CustomAction += OnCustomAction;
            mie.CustomPropertyChange += OnCustomPropertyChange;
            mie.Forward += OnForward;
            mie.Reply += OnReply;
            mie.ReplyAll += OnReplyAll;
            mie.Send += OnSend;
            mie.BeforeCheckNames += OnBeforeCheckNames;
            mie.BeforeAttachmentSave += OnBeforeAttachmentSave;
            mie.BeforeAttachmentAdd += OnBeforeAttachmentAdd;
            mie.Unload += OnUnload;
            mie.BeforeAutoSave += OnBeforeAutoSave;
            mie.BeforeRead += OnBeforeRead;

            _mailItemEventsAttached = true;
        }

        private void DetachMailItemEvents()
        {
            if (!_mailItemEventsAttached || _mailItem == null) return;
            //var mie = (MailItemEvents_Event)_mailItem;
            var mie = base._comEvents;

            mie.CustomAction -= OnCustomAction;
            mie.CustomPropertyChange -= OnCustomPropertyChange;
            mie.Forward -= OnForward;
            mie.Reply -= OnReply;
            mie.ReplyAll -= OnReplyAll;
            mie.Send -= OnSend;
            mie.BeforeCheckNames -= OnBeforeCheckNames;
            mie.BeforeAttachmentSave -= OnBeforeAttachmentSave;
            mie.BeforeAttachmentAdd -= OnBeforeAttachmentAdd;
            mie.Unload -= OnUnload;
            mie.BeforeAutoSave -= OnBeforeAutoSave;
            mie.BeforeRead -= OnBeforeRead;            

            _mailItemEventsAttached = false;
        }

        public override void Dispose()
        {
            DetachMailItemEvents();
            base.Dispose();
        }

        // --- Private handler methods to bridge events ---
        private void OnCustomAction(object Action, object Response, ref bool Cancel)
            => CustomAction?.Invoke(Action, Response, ref Cancel);

        private void OnCustomPropertyChange(string Name)
            => CustomPropertyChange?.Invoke(Name);

        //private void OnOpen(ref bool cancel) => Open?.Invoke(ref cancel);
        private void OnForward(object forward, ref bool cancel) => ForwardEvent?.Invoke(forward, ref cancel);

        private void OnReply(object Response, ref bool Cancel) => ReplyEvent?.Invoke(Response, ref Cancel);

        private void OnReplyAll(object Response, ref bool Cancel) => ReplyAllEvent?.Invoke(Response, ref Cancel);

        private void OnSend(ref bool Cancel)
            => SendEvent?.Invoke(ref Cancel);

        private void OnBeforeCheckNames(ref bool Cancel)
            => BeforeCheckNames?.Invoke(ref Cancel);

        private void OnBeforeAttachmentSave(Attachment Attachment, ref bool Cancel)
            => BeforeAttachmentSave?.Invoke(Attachment, ref Cancel);

        private void OnBeforeAttachmentAdd(Attachment Attachment, ref bool Cancel)
            => BeforeAttachmentAdd?.Invoke(Attachment, ref Cancel);

        private void OnUnload()
            => Unload?.Invoke();

        private void OnBeforeAutoSave(ref bool Cancel)
            => BeforeAutoSave?.Invoke(ref Cancel);

        private void OnBeforeRead()
            => BeforeRead?.Invoke();
    }
}
