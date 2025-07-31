using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;

namespace Gsync.OutlookInterop.Interfaces.Items
{
    public interface IItem : IEquatable<IItem>, IDisposable
    {
        // === Core Properties (common to all oltypes.core) ===
        Application Application { get; }
        OlObjectClass Class { get; }
        NameSpace Session { get; }
        object Parent { get; }
        Attachments Attachments { get; }
        string BillingInformation { get; set; }
        string Body { get; set; }
        string Categories { get; set; }
        string Companies { get; set; }
        DateTime CreationTime { get; }
        string EntryID { get; }
        OlImportance Importance { get; set; }
        object InnerObject { get; }
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

        // === Added Missing Common Properties (from oltypes.core.properties.common) ===
        Actions Actions { get; }
        string ConversationIndex { get; }
        string ConversationTopic { get; }
        string FormDescription { get; }
        object GetInspector { get; }
        object MAPIOBJECT { get; }
        object UserProperties { get; }

        // === Additional Utility Property ===
        IEqualityComparer<IItem> EqualityComparer { get; set; }

        // === Methods ===
        void Close(OlInspectorClose SaveMode);
        object Copy();
        void Delete();
        void Display(object Modal = null);
        object Move(MAPIFolder DestFldr);
        void PrintOut();
        void Save();
        void SaveAs(string Path, object Type = null);
        void ShowCategoriesDialog();

        // === Events ===
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

        // ------------------------------------------------------------------------
        //          PROPERTIES NOT PRESENT IN oltypes.core.properties.common
        // ------------------------------------------------------------------------
        /*
        string ConversationID { get; }
        string HTMLBody { get; set; }        
        ItemProperties ItemProperties { get; }
        string SenderEmailAddress { get; }
        string SenderName { get; }
        */
    }
}
