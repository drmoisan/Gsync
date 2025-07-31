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
    public class OutlookItemWrapper : IItem 
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

        #region OutlookItemWrapper

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

        #endregion OutlookItemWrapper

        #region IItem Implementation

        #region IItem Properties and Fields Implementation

        protected object _item;
        protected Type _itemType;
        protected dynamic _dyn;

        protected static readonly ImmutableHashSet<string> DefaultSupportedTypes = new HashSet<string>(
        [
            "MailItem", "TaskItem", "AppointmentItem", "ContactItem", "NoteItem",
            "JournalItem", "PostItem", "ReportItem", "DistListItem", "DocumentItem",
            "RemoteItem", "SharingItem", "StorageItem",
            "TaskRequestItem", "TaskRequestAcceptItem", "TaskRequestDeclineItem", "TaskRequestUpdateItem"
        ]).ToImmutableHashSet();
        internal virtual ImmutableHashSet<string> SupportedTypes { get; } = DefaultSupportedTypes.ToImmutableHashSet();
        
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
        public DateTime CreationTime => _dyn.CreationTime;
        public string EntryID => _dyn.EntryID;        
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

        // Add these property implementations to the OutlookItemWrapper class
        public Actions Actions => _dyn.Actions;

        public string ConversationIndex => _dyn.ConversationIndex;

        public string ConversationTopic => _dyn.ConversationTopic;

        public string FormDescription => _dyn.FormDescription;

        public object GetInspector => _dyn.GetInspector;

        public object MAPIOBJECT => _dyn.MAPIOBJECT;

        public object UserProperties => _dyn.UserProperties;

        #endregion IItem IItem Properties and Fields Implementation

        #region IItem Method Implementation

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

        #endregion IItem Method Implementation

        #region IItem Event Implementation

        #region C# Events

        public event IItem.AttachmentAddEventHandler AttachmentAdd;
        public event IItem.AttachmentReadEventHandler AttachmentRead;
        public event IItem.AttachmentRemoveEventHandler AttachmentRemove;
        public event IItem.BeforeDeleteEventHandler BeforeDelete;
        public event IItem.CloseEventHandler CloseEvent;
        public event IItem.OpenEventHandler Open;
        public event IItem.PropertyChangeEventHandler PropertyChange;
        public event IItem.ReadEventHandler Read;
        public event IItem.WriteEventHandler Write;

        #endregion C# Events

        #region COM Event Handlers => Invoke C# Events

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

        #endregion COM Event Handlers => Invoke C# Events

        #region Wire and Unwire Bridge COM Event Handlers

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

        #endregion Wire and Unwire Bridge COM Event Handlers

        #endregion IItem Event Implementation

        #region IEquatable<IItem> Implementation

        private IEqualityComparer<IItem> _equalityComparer = new IItemEqualityComparer();
        /// <summary>
        /// Gets or sets the equality comparer used for IEquatable<IItem> implementation.
        /// </summary>
        public IEqualityComparer<IItem> EqualityComparer
        {
            get => _equalityComparer;
            set => _equalityComparer = value ?? new IItemEqualityComparer();
        }

#nullable enable

        /// <summary>
        /// Implements IEquatable<IItem> using the injected or default IEqualityComparer<IItem>.
        /// </summary>
        public bool Equals(IItem? other)
        {
            return EqualityComparer.Equals(this, other);
        }

        public override bool Equals(object? obj)
        {
            if (ReferenceEquals(this, obj)) return true;
            if (obj is IItem item)
                return Equals(item);
            return false;
        }

        public override int GetHashCode()
        {
            return EqualityComparer.GetHashCode(this);
        }

#nullable disable

        #endregion IEquatable<IItem> Implementation

        #region IDisposable Implementation

        protected bool _disposed = false;

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
                
        #endregion IDisposable Implementation

        #endregion IItem Implementation
    }
}
