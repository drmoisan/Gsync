using Microsoft.Office.Interop.Outlook;
using System;
using Gsync.OutlookInterop.Interfaces.Items;
using Newtonsoft.Json;
using System.Collections.Generic;
//using System.Windows.Forms;

namespace Gsync.OutlookInterop.Item
{
    /// <summary>
    /// A detached, immutable snapshot of an Outlook item.
    /// All COM object-typed properties are always null.
    /// Provides the ability to reattach to a live Outlook item via EntryID.
    /// </summary>
    public class DetachedOutlookItem : IItem
    {
        #region ctor

        public DetachedOutlookItem() { }

        // --- Constructor: copies value properties, nullifies COM-typed properties ---
        public DetachedOutlookItem(IItem item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));

            BillingInformation = item.BillingInformation;
            Body = item.Body;
            Categories = item.Categories;
            Class = item.Class;
            Companies = item.Companies;
            CreationTime = item.CreationTime;
            EntryID = item.EntryID;
            Importance = item.Importance;
            LastModificationTime = item.LastModificationTime;
            MessageClass = item.MessageClass;
            Mileage = item.Mileage;
            NoAging = item.NoAging;
            OutlookInternalVersion = item.OutlookInternalVersion;
            OutlookVersion = item.OutlookVersion;
            Saved = item.Saved;
            Sensitivity = item.Sensitivity;
            Size = item.Size;
            Subject = item.Subject;
            UnRead = item.UnRead;
            
            ConversationIndex = item.ConversationIndex;
            ConversationTopic = item.ConversationTopic;
            FormDescription = item.FormDescription;

            var folder = item.Parent as MAPIFolder;
            StoreID = folder?.StoreID;
            ParentFolderEntryID = folder?.EntryID;

            //TODO: Move to DetachedMailItem
            //ConversationID = item.ConversationID;
            //HTMLBody = item.HTMLBody;
            //SenderEmailAddress = item.SenderEmailAddress;
            //SenderName = item.SenderName;
        }

        #endregion ctor

        #region DetachedOutlookItem Methods

        /// <summary>
        /// Attempts to reattach to a live Outlook COM item using the provided NameSpace (session).
        /// Returns a new OutlookItemWrapper or throws if not found.
        /// </summary>
        public OutlookItemWrapper Reattach(Application application)
        {
            if (application == null) throw new ArgumentNullException(nameof(application));
            if (string.IsNullOrWhiteSpace(this.EntryID) || string.IsNullOrWhiteSpace(this.StoreID))
                throw new InvalidOperationException("Insufficient information to reattach.");

            NameSpace session = application.Session;
            // Locate the right Store
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

            // Use the StoreID overload (ParentFolderEntryID optional for additional safety/context)
            object comObject = session.GetItemFromID(this.EntryID, this.StoreID);

            if (comObject == null)
                throw new InvalidOperationException("Outlook item not found in current store/session.");

            return new OutlookItemWrapper(comObject);
        }

        #endregion DetachedOutlookItem Methods

        #region IItem Implementation

        #region IItem Properties Implementation

        // --- COM reference properties: always null in detached object ---
        [JsonIgnore] public Actions Actions => null;
        [JsonIgnore] public Application Application => null;
        [JsonIgnore] public Attachments Attachments => null;
        [JsonIgnore] public object GetInspector => null;
        [JsonIgnore] public ItemProperties ItemProperties => null;
        [JsonIgnore] public NameSpace Session => null;
        [JsonIgnore] public object InnerObject => null;
        [JsonIgnore] public object MAPIOBJECT => null;
        [JsonIgnore] public object Parent => null;
        [JsonIgnore] public object UserProperties => null;

        // --- Value & string properties ---
        public string BillingInformation { get; set; }
        public string Body { get; set; }
        public string Categories { get; set; }
        public OlObjectClass Class { get; set; }
        public string Companies { get; set; }
        public string ConversationID { get; set; }
        public string ConversationIndex { get; set; }
        public string ConversationTopic { get; set; }
        public DateTime CreationTime { get; set; }
        public string EntryID { get; set; }
        public string FormDescription { get; set; }
        public string HTMLBody { get; set; }
        public OlImportance Importance { get; set; }
        public DateTime LastModificationTime { get; set; }
        public string MessageClass { get; set; }
        public string Mileage { get; set; }
        public bool NoAging { get; set; }
        public int OutlookInternalVersion { get; set; }
        public string OutlookVersion { get; set; }
        public bool Saved { get; set; }
        public string SenderEmailAddress { get; set; }
        public string SenderName { get; set; }
        public OlSensitivity Sensitivity { get; set; }
        public int Size { get; set; }
        public string Subject { get; set; }
        public bool UnRead { get; set; }

        // --- Optionally store parent folder ID for session-scoped reattachment ---
        public string ParentFolderEntryID { get; set; }

        public string StoreID { get; set; }

        #endregion IItem Properties Implementation

        #region IItem Methods Implementation

        // --- Methods: always throw NotSupportedException ---
        public void Close(OlInspectorClose SaveMode) =>
            throw new NotSupportedException("This DetachedOutlookItem is not connected to Outlook.");

        public object Copy() =>
            throw new NotSupportedException("This DetachedOutlookItem is not connected to Outlook.");

        public void Delete() =>
            throw new NotSupportedException("This DetachedOutlookItem is not connected to Outlook.");

        public void Display(object Modal = null) =>
            throw new NotSupportedException("This DetachedOutlookItem is not connected to Outlook.");

        public object Move(MAPIFolder DestFldr) =>
            throw new NotSupportedException("This DetachedOutlookItem is not connected to Outlook.");

        public void PrintOut() =>
            throw new NotSupportedException("This DetachedOutlookItem is not connected to Outlook.");

        public void Save() =>
            throw new NotSupportedException("This DetachedOutlookItem is not connected to Outlook.");

        public void SaveAs(string Path, object Type = null) =>
            throw new NotSupportedException("This DetachedOutlookItem is not connected to Outlook.");

        public void ShowCategoriesDialog() =>
            throw new NotSupportedException("This DetachedOutlookItem is not connected to Outlook.");

        #endregion IItem Methods Implementation

        #region IItem Events Implementation

        // --- Events: No-ops. These events can never be raised on a detached object. ---        
        public event IItem.AttachmentAddEventHandler AttachmentAdd { add { } remove { } }
        public event IItem.AttachmentReadEventHandler AttachmentRead { add { } remove { } }
        public event IItem.AttachmentRemoveEventHandler AttachmentRemove { add { } remove { } }
        public event IItem.BeforeDeleteEventHandler BeforeDelete { add { } remove { } }
        public event IItem.CloseEventHandler CloseEvent { add { } remove { } }
        public event IItem.OpenEventHandler Open { add { } remove { } }
        public event IItem.PropertyChangeEventHandler PropertyChange { add { } remove { } }
        public event IItem.ReadEventHandler Read { add { } remove { } }
        public event IItem.WriteEventHandler Write { add { } remove { } }

        #endregion IItem Events Implementation

        #region IEquatable<IItem> implementation

        private IEqualityComparer<IItem> _equalityComparer = new IItemEqualityComparer();
        /// <summary>
        /// Gets or sets the equality comparer used for IEquatable<IItem> implementation.
        /// </summary>
        [JsonIgnore]
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

        #endregion IEquatable<IItem> implementation

        #region IDisposable Implementation

        protected bool _disposed = false;

        public virtual void Dispose()
        {
            if (_disposed) return;
            //DetachComEvents();
            //ReleaseComObject(_dyn);
            //ReleaseComObject(_item);
            //ReleaseComObject(_comEvents);

            _disposed = true;
            GC.SuppressFinalize(this);
        }

        #endregion IDisposable Implementation

        #endregion IItem Implementation

    }
}
