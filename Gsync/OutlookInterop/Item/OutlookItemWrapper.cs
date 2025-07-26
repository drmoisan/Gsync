using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.OutlookInterop.Item
{
    using Gsync.OutlookInterop.Interfaces.Items;
    using Microsoft.Office.Interop.Outlook;
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Runtime.InteropServices;

    public class OutlookItemWrapper : IItem, IDisposable
    {
        private readonly object _item;
        private readonly Type _itemType;
        private bool _disposed = false;

        // Cache for PropertyInfo and MethodInfo
        private static readonly ConcurrentDictionary<Type, Dictionary<string, PropertyInfo>> PropertyCache = new();
        private static readonly ConcurrentDictionary<Type, Dictionary<string, MethodInfo>> MethodCache = new();

        // Supported Outlook item types (COM class names)
        private static readonly HashSet<string> SupportedTypes = new(
            new[]
            {
            "MailItem", "TaskItem", "AppointmentItem", "ContactItem", "NoteItem",
            "JournalItem", "PostItem", "ReportItem", "DistListItem", "DocumentItem",
            "RemoteItem", "SharingItem", "StorageItem",
            "TaskRequestItem", "TaskRequestAcceptItem", "TaskRequestDeclineItem", "TaskRequestUpdateItem"
            }
        );

        // Event bridging
        private ItemEvents_10_Event _comEvents;
        private readonly List<Delegate> _eventHandlers = new();

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

        public OutlookItemWrapper(object item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));

            var type = OutlookItemType.GetType(item);
            string name = type.Name;

            if (!SupportedTypes.Contains(name))
                throw new ArgumentException($"Object type '{name}' is not a supported Outlook item type.");

            _item = item;
            _itemType = type;

            // Set up COM event bridge
            AttachComEvents();
        }
        
        #region Forward Properties

        // Forward properties (expand as needed)
        public Application Application => (Application)_itemType.GetProperty("Application")?.GetValue(_item);
        public Attachments Attachments => (Attachments)_itemType.GetProperty("Attachments")?.GetValue(_item);

        public string BillingInformation
        {
            get => (string)_itemType.GetProperty("BillingInformation")?.GetValue(_item);
            set => _itemType.GetProperty("BillingInformation")?.SetValue(_item, value);
        }

        public string Body
        {
            get => (string)_itemType.GetProperty("Body")?.GetValue(_item);
            set => _itemType.GetProperty("Body")?.SetValue(_item, value);
        }

        public string Categories
        {
            get => (string)_itemType.GetProperty("Categories")?.GetValue(_item);
            set => _itemType.GetProperty("Categories")?.SetValue(_item, value);
        }

        public OlObjectClass Class => (OlObjectClass)_itemType.GetProperty("Class")?.GetValue(_item);
        public string Companies
        {
            get => (string)_itemType.GetProperty("Companies")?.GetValue(_item);
            set => _itemType.GetProperty("Companies")?.SetValue(_item, value);
        }
        public string ConversationID => (string)_itemType.GetProperty("ConversationID")?.GetValue(_item);
        public DateTime CreationTime => (DateTime)_itemType.GetProperty("CreationTime")?.GetValue(_item);
        public string EntryID => (string)_itemType.GetProperty("EntryID")?.GetValue(_item);

        public string HTMLBody
        {
            get => (string)_itemType.GetProperty("HTMLBody")?.GetValue(_item);
            set => _itemType.GetProperty("HTMLBody")?.SetValue(_item, value);
        }

        public OlImportance Importance
        {
            get => (OlImportance)_itemType.GetProperty("Importance")?.GetValue(_item);
            set => _itemType.GetProperty("Importance")?.SetValue(_item, value);
        }

        public ItemProperties ItemProperties => (ItemProperties)_itemType.GetProperty("ItemProperties")?.GetValue(_item);
        public DateTime LastModificationTime => (DateTime)_itemType.GetProperty("LastModificationTime")?.GetValue(_item);
        public string MessageClass => (string)_itemType.GetProperty("MessageClass")?.GetValue(_item);

        public string Mileage
        {
            get => (string)_itemType.GetProperty("Mileage")?.GetValue(_item);
            set => _itemType.GetProperty("Mileage")?.SetValue(_item, value);
        }

        public bool NoAging
        {
            get => (bool)_itemType.GetProperty("NoAging")?.GetValue(_item);
            set => _itemType.GetProperty("NoAging")?.SetValue(_item, value);
        }

        public int OutlookInternalVersion => (int)_itemType.GetProperty("OutlookInternalVersion")?.GetValue(_item);
        public string OutlookVersion => (string)_itemType.GetProperty("OutlookVersion")?.GetValue(_item);
        public object Parent => _itemType.GetProperty("Parent")?.GetValue(_item);
        public bool Saved => (bool)_itemType.GetProperty("Saved")?.GetValue(_item);
        public string SenderEmailAddress => (string)_itemType.GetProperty("SenderEmailAddress")?.GetValue(_item);
        public string SenderName => (string)_itemType.GetProperty("SenderName")?.GetValue(_item);

        public OlSensitivity Sensitivity
        {
            get => (OlSensitivity)_itemType.GetProperty("Sensitivity")?.GetValue(_item);
            set => _itemType.GetProperty("Sensitivity")?.SetValue(_item, value);
        }

        public NameSpace Session => (NameSpace)_itemType.GetProperty("Session")?.GetValue(_item);
        public int Size => (int)_itemType.GetProperty("Size")?.GetValue(_item);

        public string Subject
        {
            get => (string)_itemType.GetProperty("Subject")?.GetValue(_item);
            set => _itemType.GetProperty("Subject")?.SetValue(_item, value);
        }

        public bool UnRead
        {
            get => (bool)_itemType.GetProperty("UnRead")?.GetValue(_item);
            set => _itemType.GetProperty("UnRead")?.SetValue(_item, value);
        }

        #endregion Forward Properties

        #region Methods

        // Methods
        public void Close(OlInspectorClose SaveMode) => _itemType.GetMethod("Close")?.Invoke(_item, new object[] { SaveMode });
        public object Copy() => _itemType.GetMethod("Copy")?.Invoke(_item, null);
        public void Delete() => _itemType.GetMethod("Delete")?.Invoke(_item, null);

        public void Display(object Modal = null) =>
            _itemType.GetMethod("Display")?.Invoke(_item, Modal != null ? new[] { Modal } : null);

        public object Move(MAPIFolder DestFldr) => _itemType.GetMethod("Move")?.Invoke(_item, new object[] { DestFldr });
        public void PrintOut() => _itemType.GetMethod("PrintOut")?.Invoke(_item, null);
        public void Save() => _itemType.GetMethod("Save")?.Invoke(_item, null);

        public void SaveAs(string Path, object Type = null) =>
            _itemType.GetMethod("SaveAs")?.Invoke(_item, Type != null ? new object[] { Path, Type } : new object[] { Path });

        public void ShowCategoriesDialog() => _itemType.GetMethod("ShowCategoriesDialog")?.Invoke(_item, null);

        // --- Event Bridging ---

        private void AttachComEvents()
        {
            // This casts to the ItemEvents_10_Event interface generated by interop
            _comEvents = _item as ItemEvents_10_Event;
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

        // --- IDisposable support for event cleanup ---
        public void Dispose()
        {
            if (_disposed) return;
            DetachComEvents();

            // Explicitly release COM object(s)
            if (_item != null && Marshal.IsComObject(_item))
            {
                try { Marshal.ReleaseComObject(_item); } catch { }
            }
            if (_comEvents != null && Marshal.IsComObject(_comEvents))
            {
                try { Marshal.ReleaseComObject(_comEvents); } catch { }
            }
            _disposed = true;
            GC.SuppressFinalize(this);
        }

        #endregion Methods

        // Delegates for event bridging (must match ItemEvents_10_Event signatures)
        //public delegate void AttachmentAddEventHandler(Attachment attachment);
        //public delegate void AttachmentReadEventHandler(Attachment attachment);
        //public delegate void AttachmentRemoveEventHandler(Attachment attachment);
        //public delegate void BeforeDeleteEventHandler(object item, ref bool cancel);
        //public delegate void CloseEventHandler();
        //public delegate void OpenEventHandler(ref bool cancel);
        //public delegate void PropertyChangeEventHandler(string name);
        //public delegate void ReadEventHandler();
        //public delegate void WriteEventHandler(ref bool cancel);
    }

}
