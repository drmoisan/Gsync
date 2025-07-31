using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.OutlookInterop.Interfaces.Items
{
    public interface IItem: IEquatable<IItem>, IDisposable
    {
        // Properties
        Application Application { get; }
        Attachments Attachments { get; }
        string BillingInformation { get; set; }
        string Body { get; set; }
        string Categories { get; set; }
        OlObjectClass Class { get; }
        string Companies { get; set; }
        string ConversationID { get; }
        DateTime CreationTime { get; }
        string EntryID { get; }
        string HTMLBody { get; set; }
        OlImportance Importance { get; set; }
        object InnerObject { get; }
        ItemProperties ItemProperties { get; }
        DateTime LastModificationTime { get; }
        string MessageClass { get; }
        string Mileage { get; set; }
        bool NoAging { get; set; }
        int OutlookInternalVersion { get; }
        string OutlookVersion { get; }
        object Parent { get; }
        bool Saved { get; }
        string SenderEmailAddress { get; }
        string SenderName { get; }
        OlSensitivity Sensitivity { get; set; }
        NameSpace Session { get; }
        int Size { get; }
        string Subject { get; set; }
        bool UnRead { get; set; }
        
        // Methods
        void Close(OlInspectorClose SaveMode);
        object Copy();
        void Delete();
        void Display(object Modal = null);
        object Move(MAPIFolder DestFldr);
        void PrintOut();
        void Save();
        void SaveAs(string Path, object Type = null);
        void ShowCategoriesDialog();
        // Events
        //event ItemEvents_10_AttachmentAddEventHandler AttachmentAdd;
        //event ItemEvents_10_AttachmentReadEventHandler AttachmentRead;
        //event ItemEvents_10_AttachmentRemoveEventHandler AttachmentRemove;
        //event ItemEvents_10_BeforeDeleteEventHandler BeforeDelete;
        //event ItemEvents_10_CloseEventHandler CloseEvent;
        //event ItemEvents_10_OpenEventHandler Open;
        //event ItemEvents_10_PropertyChangeEventHandler PropertyChange;
        //event ItemEvents_10_ReadEventHandler Read;
        //event ItemEvents_10_WriteEventHandler Write;
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

        IEqualityComparer<IItem> EqualityComparer { get; set; }
    }
}
