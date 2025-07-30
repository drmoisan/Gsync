using Gsync.OutlookInterop.Interfaces.Items;
using Gsync.Utilities.Extensions;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Gsync.OutlookInterop.Item
{
    public class OutlookItemWrapper : IItem, IDisposable
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            MethodBase.GetCurrentMethod().DeclaringType);

        #region ctor

        public OutlookItemWrapper(object item)
            : this(item, item as ItemEvents_10_Event)
        {
            Init();
        }

        protected OutlookItemWrapper(object item, ItemEvents_10_Event comEvents)
        {
            _item = item;
            _comEvents = comEvents;
        }

        protected OutlookItemWrapper(object item, ItemEvents_10_Event comEvents, ImmutableHashSet<string> supportedTypes)
        {
            _item = item;
            _comEvents = comEvents;
            SupportedTypes = supportedTypes;
        }

        protected OutlookItemWrapper Init()
        {
            _item.ThrowIfNull();

            var type = OutlookItemType.GetType(_item);
            string name = type?.Name ?? $"Unknown Type";

            if (!SupportedTypes.Contains(name))
                throw new ArgumentException($"Object type '{name}' is not a supported Outlook item type.");

            _comEvents.ThrowIfNull();

            _dyn = _item;
            _itemType = type;

            AttachComEvents();
            return this;
        }

        #endregion ctor

        #region Private Fields and Properties

        protected object _item;
        protected Type _itemType;
        protected dynamic _dyn;
        protected bool _disposed = false;

        private static readonly ConcurrentDictionary<Type, Dictionary<string, PropertyInfo>> PropertyCache = new();
        private static readonly ConcurrentDictionary<Type, Dictionary<string, MethodInfo>> MethodCache = new();

        private static readonly ImmutableHashSet<string> DefaultSupportedTypes = new HashSet<string>(
        [
            "MailItem", "TaskItem", "AppointmentItem", "ContactItem", "NoteItem",
            "JournalItem", "PostItem", "ReportItem", "DistListItem", "DocumentItem",
            "RemoteItem", "SharingItem", "StorageItem",
            "TaskRequestItem", "TaskRequestAcceptItem", "TaskRequestDeclineItem", "TaskRequestUpdateItem"
        ]).ToImmutableHashSet();

        internal virtual ImmutableHashSet<string> SupportedTypes { get; } = DefaultSupportedTypes.ToImmutableHashSet();

        #endregion Private Fields and Properties

        #region Properties

        public Application Application => _dyn.Application;
        public Attachments Attachments => _dyn.Attachments;
        public string BillingInformation
        {
            get => _dyn.BillingInformation;
            set => _dyn.BillingInformation = value;
        }
        public string Body
        {
            get => _dyn.Body;
            set => _dyn.Body = value;
        }
        public string Categories
        {
            get => _dyn.Categories;
            set => _dyn.Categories = value;
        }
        public OlObjectClass Class => _dyn.Class;
        public string Companies
        {
            get => _dyn.Companies;
            set => _dyn.Companies = value;
        }
        public string ConversationID => _dyn.ConversationID;
        public DateTime CreationTime => _dyn.CreationTime;
        public string EntryID => _dyn.EntryID;
        public string HTMLBody
        {
            get => _dyn.HTMLBody;
            set => _dyn.HTMLBody = value;
        }
        public OlImportance Importance
        {
            get => _dyn.Importance;
            set => _dyn.Importance = value;
        }
        public object InnerObject => _item;
        public ItemProperties ItemProperties => _dyn.ItemProperties;
        public DateTime LastModificationTime => _dyn.LastModificationTime;
        public string MessageClass => _dyn.MessageClass;
        public string Mileage
        {
            get => _dyn.Mileage;
            set => _dyn.Mileage = value;
        }
        public bool NoAging
        {
            get => _dyn.NoAging;
            set => _dyn.NoAging = value;
        }
        public int OutlookInternalVersion => _dyn.OutlookInternalVersion;
        public string OutlookVersion => _dyn.OutlookVersion;
        public object Parent => _dyn.Parent;
        public bool Saved => _dyn.Saved;
        public string SenderEmailAddress => _dyn.SenderEmailAddress;
        public string SenderName => _dyn.SenderName;
        public OlSensitivity Sensitivity
        {
            get => _dyn.Sensitivity;
            set => _dyn.Sensitivity = value;
        }
        public NameSpace Session => _dyn.Session;
        public int Size => _dyn.Size;
        public string Subject
        {
            get => _dyn.Subject;
            set => _dyn.Subject = value;
        }
        public bool UnRead
        {
            get => _dyn.UnRead;
            set => _dyn.UnRead = value;
        }
        #endregion Properties

        #region Public Methods

        public void Close(OlInspectorClose SaveMode)
        {
            var type = _item.ThrowIfNull("Cannot close item because it is null").GetType();
            var method = _item.GetType().GetMethod("Close", new[] { typeof(OlInspectorClose) });            
            if (method is null) { throw new InvalidOperationException($"Method 'Close' not found on type '{type.FullName}'."); }
            method.Invoke(_item, new object[] { SaveMode });
        }
        public object Copy()
        {
            return _dyn.Copy();
        }
        public void Delete()
        {
            _dyn.Delete();
        }
        public void Display(object Modal = null)
        {
            if (Modal != null) _dyn.Display(Modal);
            else _dyn.Display();
        }
        public object Move(MAPIFolder DestFldr)
        {
            return _dyn.Move(DestFldr);
        }
        public void PrintOut()
        {
            _dyn.PrintOut();
        }
        public void Save()
        {
            _dyn.Save();
        }
        public void SaveAs(string Path, object Type = null)
        {
            if (Type != null) _dyn.SaveAs(Path, Type);
            else _dyn.SaveAs(Path);
        }
        public void ShowCategoriesDialog()
        {
            _dyn.ShowCategoriesDialog();
        }

        public virtual void Dispose()
        {
            if (_disposed) return;
            DetachComEvents();
            ReleaseComObject(_dyn);
            ReleaseComObject(_item);
            ReleaseComObject(_comEvents);

            _disposed = true;
            GC.SuppressFinalize(this);
        }

        #endregion Public Methods

        #region Event Bridging

        public event IItem.AttachmentAddEventHandler AttachmentAdd;
        public event IItem.AttachmentReadEventHandler AttachmentRead;
        public event IItem.AttachmentRemoveEventHandler AttachmentRemove;
        public event IItem.BeforeDeleteEventHandler BeforeDelete;
        public event IItem.CloseEventHandler CloseEvent;
        public event IItem.OpenEventHandler Open;
        public event IItem.PropertyChangeEventHandler PropertyChange;
        public event IItem.ReadEventHandler Read;
        public event IItem.WriteEventHandler Write;

        private readonly List<Delegate> _eventHandlers = new();

        private void OnAttachmentAdd(Attachment attachment) => AttachmentAdd?.Invoke(attachment);
        private void OnAttachmentRead(Attachment attachment) => AttachmentRead?.Invoke(attachment);
        private void OnAttachmentRemove(Attachment attachment) => AttachmentRemove?.Invoke(attachment);
        private void OnBeforeDelete(object item, ref bool cancel) => BeforeDelete?.Invoke(item, ref cancel);
        private void OnCloseEvent(ref bool cancel) => CloseEvent?.Invoke(ref cancel);
        private void OnOpen(ref bool cancel) => Open?.Invoke(ref cancel);
        private void OnPropertyChange(string name) => PropertyChange?.Invoke(name);
        private void OnRead() => Read?.Invoke();
        private void OnWrite(ref bool cancel) => Write?.Invoke(ref cancel);

        protected ItemEvents_10_Event _comEvents;

        private void AttachComEvents()
        {
            if (_comEvents == null) return;

            _comEvents.AttachmentAdd += OnAttachmentAdd;
            _comEvents.AttachmentRead += OnAttachmentRead;
            _comEvents.AttachmentRemove += OnAttachmentRemove;
            _comEvents.BeforeDelete += OnBeforeDelete;
            _comEvents.Close += OnCloseEvent;
            _comEvents.Open += OnOpen;
            _comEvents.PropertyChange += OnPropertyChange;
            _comEvents.Read += OnRead;
            _comEvents.Write += OnWrite;
        }

        private void DetachComEvents()
        {
            if (_comEvents == null) return;

            _comEvents.AttachmentAdd -= OnAttachmentAdd;
            _comEvents.AttachmentRead -= OnAttachmentRead;
            _comEvents.AttachmentRemove -= OnAttachmentRemove;
            _comEvents.BeforeDelete -= OnBeforeDelete;
            _comEvents.Close -= OnCloseEvent;
            _comEvents.Open -= OnOpen;
            _comEvents.PropertyChange -= OnPropertyChange;
            _comEvents.Read -= OnRead;
            _comEvents.Write -= OnWrite;
        }

        #endregion Event Bridging

        #region Private Helper Methods

        protected virtual bool IsComObjectFunc(object obj)
        {
            if (obj is null) { return false; }
            else { return Marshal.IsComObject(obj); }
        }
        protected virtual void ReleaseComObject(object comObj)
        {
            if (comObj != null && IsComObjectFunc(comObj))
                Marshal.ReleaseComObject(comObj);
        }

        #endregion Private Helper Methods

    }
}
