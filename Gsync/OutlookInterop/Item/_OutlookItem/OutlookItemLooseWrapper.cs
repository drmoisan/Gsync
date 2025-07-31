using Gsync.OutlookInterop.Interfaces.Items;
using Gsync.Utilities.Extensions;
using log4net.Repository.Hierarchy;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Gsync.OutlookInterop.Item
{
    public class OutlookItemLooseWrapper : IItem, IDisposable, IEquatable<IItem>
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #region ctor

        public OutlookItemLooseWrapper(object item)
            : this(item, item as ItemEvents_10_Event)
        {
            Init();
        }

        protected OutlookItemLooseWrapper(object item, ItemEvents_10_Event comEvents)
        {
            _item = item;
            _comEvents = comEvents;
        }

        protected OutlookItemLooseWrapper(object item, ItemEvents_10_Event comEvents, ImmutableHashSet<string> supportedTypes)

        {
            _item = item;
            _comEvents = comEvents;
            SupportedTypes = supportedTypes;
        }

        protected OutlookItemLooseWrapper Init()
        {
            // Initialize any additional properties or state here if needed
            _item.ThrowIfNull();

            var type = OutlookItemType.GetType(_item);
            string name = type?.Name ?? $"Unknown Type";

            if (!SupportedTypes.Contains(name))
                throw new ArgumentException($"Object type '{name}' is not a supported Outlook item type.");

            // This should never be null if the item is supported
            _comEvents.ThrowIfNull();

            _dyn = _item;
            _itemType = type;

            // Set up COM event bridge

            AttachComEvents();
            return this;
        }

        #endregion ctor

        #region OutlookItemLooseWrapper

        #region Private Fields and Properties

        private readonly object _item;
        private Type _itemType;
        private dynamic _dyn;
        private bool _disposed = false;

        // Cache for PropertyInfo and MethodInfo
        private static readonly ConcurrentDictionary<Type, Dictionary<string, PropertyInfo>> PropertyCache = new();
        private static readonly ConcurrentDictionary<Type, Dictionary<string, MethodInfo>> MethodCache = new();

        // Supported Outlook item types (COM class names)        
        private static readonly ImmutableHashSet<string> DefaultSupportedTypes = new HashSet<string>(
        [
            "MailItem", "TaskItem", "AppointmentItem", "ContactItem", "NoteItem",
            "JournalItem", "PostItem", "ReportItem", "DistListItem", "DocumentItem",
            "RemoteItem", "SharingItem", "StorageItem",
            "TaskRequestItem", "TaskRequestAcceptItem", "TaskRequestDeclineItem", "TaskRequestUpdateItem"
        ]).ToImmutableHashSet();

        #endregion Private Fields and Properties

        #region Private Helper Methods

        // --- Helper safe wrappers ---
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
        private T TryGet<T>(Func<T> getter)
        {
            try { return getter(); }
            catch (System.Exception ex)
            {
                logger.Error($"Exception in TryGet<{typeof(T).Name}>: {ex.Message}", ex);
                return default;
            }
        }
        private void TrySet(System.Action setter)
        {
            try { setter(); }
            catch (System.Exception ex)
            {
                logger.Error($"Exception in TrySet: {ex.Message}", ex);
            }
        }

        #endregion Private Helper Methods

        #endregion OutlookItemLooseWrapper

        #region IItem Implementation

        #region IItem Properties Implementation

        // Forward properties (expand as needed)
        public Application Application => TryGet(() => (Application)_dyn.Application);
        public Attachments Attachments => TryGet(() => (Attachments)_dyn.Attachments);
        public string BillingInformation
        {
            get => TryGet(() => (string)_dyn.BillingInformation);
            set => TrySet(() => _dyn.BillingInformation = value);
        }
        public string Body
        {
            get => TryGet(() => (string)_dyn.Body);
            set => TrySet(() => _dyn.Body = value);
        }
        public string Categories
        {
            get => TryGet(() => (string)_dyn.Categories);
            set => TrySet(() => _dyn.Categories = value);
        }
        public OlObjectClass Class => TryGet(() => (OlObjectClass)_dyn.Class);
        public string Companies
        {
            get => TryGet(() => (string)_dyn.Companies);
            set => TrySet(() => _dyn.Companies = value);
        }
        public string ConversationID => TryGet(() => (string)_dyn.ConversationID);
        public DateTime CreationTime => TryGet(() => (DateTime)_dyn.CreationTime);
        public string EntryID => TryGet(() => (string)_dyn.EntryID);
        public string HTMLBody
        {
            get => TryGet(() => (string)_dyn.HTMLBody);
            set => TrySet(() => _dyn.HTMLBody = value);
        }
        public OlImportance Importance
        {
            get => TryGet(() => (OlImportance)_dyn.Importance);
            set => TrySet(() => _dyn.Importance = value);
        }
        public object InnerObject => TryGet(() => (object)_dyn.InnerObject);
        public ItemProperties ItemProperties => TryGet(() => (ItemProperties)_dyn.ItemProperties);
        public DateTime LastModificationTime => TryGet(() => (DateTime)_dyn.LastModificationTime);
        public string MessageClass => TryGet(() => (string)_dyn.MessageClass);
        public string Mileage
        {
            get => TryGet(() => (string)_dyn.Mileage);
            set => TrySet(() => _dyn.Mileage = value);
        }
        public bool NoAging
        {
            get => TryGet(() => (bool)_dyn.NoAging);
            set => TrySet(() => _dyn.NoAging = value);
        }
        public int OutlookInternalVersion => TryGet(() => (int)_dyn.OutlookInternalVersion);
        public string OutlookVersion => TryGet(() => (string)_dyn.OutlookVersion);
        public object Parent => TryGet(() => (object)_dyn.Parent);
        public bool Saved => TryGet(() => (bool)_dyn.Saved);
        public string SenderEmailAddress => TryGet(() => (string)_dyn.SenderEmailAddress);
        public string SenderName => TryGet(() => (string)_dyn.SenderName);
        public OlSensitivity Sensitivity
        {
            get => TryGet(() => (OlSensitivity)_dyn.Sensitivity);
            set => TrySet(() => _dyn.Sensitivity = value);
        }
        public NameSpace Session => TryGet(() => (NameSpace)_dyn.Session);
        public int Size => TryGet(() => (int)_dyn.Size);
        public string Subject
        {
            get => TryGet(() => (string)_dyn.Subject);
            set => TrySet(() => _dyn.Subject = value);
        }
        public bool UnRead
        {
            get => TryGet(() => (bool)_dyn.UnRead);
            set => TrySet(() => _dyn.UnRead = value);
        }
        
        #endregion IItem Properties Implementation

        #region IItem Method Implementation

        // Methods        
        public void Close(OlInspectorClose SaveMode)
        {
            try
            {
                var method = _item.GetType().GetMethod("Close", new[] { typeof(OlInspectorClose) });
                method.Invoke(_item, new object[] { SaveMode });
            }
            catch (System.Exception e)
            {
                logger.Error($"Error closing item: {e.Message}", e);
            }
        }
        public object Copy()
        {
            return TryGet(() => _dyn.Copy());
        }
        public void Delete()
        {
            TrySet(() => _dyn.Delete());
        }
        public void Display(object Modal = null)
        {
            TrySet(() => { if (Modal != null) _dyn.Display(Modal); else _dyn.Display(); });
        }
        public object Move(MAPIFolder DestFldr)
        {
            return TryGet(() => _dyn.Move(DestFldr));
        }
        public void PrintOut()
        {
            TrySet(() => _dyn.PrintOut());
        }
        public void Save()
        {
            TrySet(() => _dyn.Save());
        }
        public void SaveAs(string Path, object Type = null)
        {
            TrySet(() => { if (Type != null) _dyn.SaveAs(Path, Type); else _dyn.SaveAs(Path); });
        }
        public void ShowCategoriesDialog()
        {
            TrySet(() => _dyn.ShowCategoriesDialog());
        }

        #endregion IItem Method Implementation
        
        #region IItem Event Implementation

        #region C# Events

        // C# Events (bridged from COM)
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

        // Holder for COM event handlers that invoke C# events
        private readonly List<Delegate> _eventHandlers = new();
        // COM event handlers that raise .NET events
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

        private ItemEvents_10_Event _comEvents;
        private void AttachComEvents()
        {
            if (_comEvents == null) return;

            // Wire each COM event to .NET event
            if (_comEvents != null)
            {
                _comEvents.AttachmentAdd += OnAttachmentAdd;
                _comEvents.AttachmentRead += OnAttachmentRead;
                _comEvents.AttachmentRemove += OnAttachmentRemove;
                _comEvents.BeforeDelete += OnBeforeDelete;
                _comEvents.Close += OnCloseEvent;
                _comEvents.Open += OnOpen;
                _comEvents.PropertyChange += OnPropertyChange;
                _comEvents.Read += OnRead;
                _comEvents.Write += OnWrite;

                // Store for cleanup
                _eventHandlers.Add((IItem.AttachmentAddEventHandler)OnAttachmentAdd);
                _eventHandlers.Add((IItem.AttachmentReadEventHandler)OnAttachmentRead);
                _eventHandlers.Add((IItem.AttachmentRemoveEventHandler)OnAttachmentRemove);
                _eventHandlers.Add((IItem.BeforeDeleteEventHandler)OnBeforeDelete);
                _eventHandlers.Add((IItem.CloseEventHandler)OnCloseEvent);
                _eventHandlers.Add((IItem.OpenEventHandler)OnOpen);
                _eventHandlers.Add((IItem.PropertyChangeEventHandler)OnPropertyChange);
                _eventHandlers.Add((IItem.ReadEventHandler)OnRead);
                _eventHandlers.Add((IItem.WriteEventHandler)OnWrite);
            }
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

        #region IDisposable Implementation

        // --- IDisposable support for event cleanup ---
        public void Dispose()
        {
            if (_disposed) return;
            DetachComEvents();

            // Explicitly release COM object(s)
            ReleaseComObject(_dyn);
            ReleaseComObject(_item);
            ReleaseComObject(_comEvents);

            _disposed = true;
            GC.SuppressFinalize(this);
        }

        #endregion IDisposable Implementation

        #region IEquatable<IItem> Implementation

        internal virtual ImmutableHashSet<string> SupportedTypes { get; } = DefaultSupportedTypes.ToImmutableHashSet();

        // --- Equality Comparer for IEquatable<IItem> ---
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

        #endregion IItem Implementation
    }
}
