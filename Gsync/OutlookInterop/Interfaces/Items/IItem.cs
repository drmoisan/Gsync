using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;

namespace Gsync.OutlookInterop.Interfaces.Items
{
    public interface IItem : IEquatable<IItem>, IDisposable
    {
        #region Custom IItem Properties and Methods

        object InnerObject { get; }

        string RawHeaders { get; }

        string MessageId { get; }

        #endregion Custom IItem Properties and Methods

        #region CORE - Assembly Microsoft.Office.Interop.Outlook, Version=15.0.0.0

        #region Core Properties (common to all oltypes)

        Application Application { get; }
        OlObjectClass Class { get; }
        NameSpace Session { get; }
        object Parent { get; }
        // TODO: Write a wrapper for Outlook.Attachments
        Attachments Attachments { get; }
        string BillingInformation { get; set; }
        string Body { get; set; }
        string Categories { get; set; }
        string Companies { get; set; }
        string ConversationID { get; }
        DateTime CreationTime { get; }
        string EntryID { get; }
        OlImportance Importance { get; set; }        
        DateTime LastModificationTime { get; }
        string MessageClass { get; }
        string Mileage { get; set; }
        bool NoAging { get; set; }
        int OutlookInternalVersion { get; }
        string OutlookVersion { get; }
        bool Saved { get; }
        OlSensitivity Sensitivity { get; set; }
        int Size { get; }
        string Subject { get; set; }
        bool UnRead { get; set; }
        // TODO: Write a wrapper for Outlook.Actions
        Actions Actions { get; }
        string ConversationIndex { get; }
        string ConversationTopic { get; }
        string FormDescription { get; }
        object GetInspector { get; }
        // TODO: Write a wrapper for MAPIOBJECT
        object MAPIOBJECT { get; }
        object UserProperties { get; }
        bool AutoResolvedWinner { get; }
        Conflicts Conflicts { get; }
        OlDownloadState DownloadState { get; }
        bool IsConflict { get; }
        Links Links { get; }
        PropertyAccessor PropertyAccessor { get; }        

        #endregion Core Properties (common to all oltypes)

        #region Core Methods (common to all oltypes)

        // TODO: Change Close signature to not rely on Outlook constants
        void Close(OlInspectorClose SaveMode);
        object Copy();
        void Delete();
        void Display(object Modal = null);
        // TODO: Change Move signature to not rely on Outlook constants
        object Move(MAPIFolder DestFldr);
        void PrintOut();
        void Save();
        void SaveAs(string Path, object Type = null);
        
        #endregion Core Methods (common to all oltypes)

        #region Core Events (common to all oltypes)
        
        event AttachmentAddEventHandler AttachmentAdd;
        event AttachmentReadEventHandler AttachmentRead;
        event AttachmentRemoveEventHandler AttachmentRemove;
        event BeforeDeleteEventHandler BeforeDelete;
        event CloseEventHandler CloseEvent;
        event OpenEventHandler Open;
        event PropertyChangeEventHandler PropertyChange;
        event ReadEventHandler Read;
        event WriteEventHandler Write;

        delegate void AttachmentAddEventHandler(Attachment attachment);
        delegate void AttachmentReadEventHandler(Attachment attachment);
        delegate void AttachmentRemoveEventHandler(Attachment attachment);
        delegate void BeforeDeleteEventHandler(object item, ref bool cancel);
        delegate void CloseEventHandler(ref bool cancel);
        delegate void OpenEventHandler(ref bool cancel);
        delegate void PropertyChangeEventHandler(string name);
        delegate void ReadEventHandler();
        delegate void WriteEventHandler(ref bool cancel);
        
        #endregion Core Events (common to all oltypes)

        #endregion CORE - Assembly Microsoft.Office.Interop.Outlook, Version=15.0.0.0

        #region IEquatable<IItem> Implementation

        // === Additional Utility Property ===
        IEqualityComparer<IItem> EqualityComparer { get; set; }

        #endregion IEquatable<IItem> Implementation
        
    }
}
