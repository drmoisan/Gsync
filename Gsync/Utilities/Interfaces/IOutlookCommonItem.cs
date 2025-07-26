using Microsoft.Office.Interop.Outlook;
using System;

namespace Gsync.Utilities.Interfaces
{
    public interface IOutlookCommonItem
    {
        // Common Properties
        string EntryID { get; }
        string Subject { get; set; }
        string Body { get; set; }
        string MessageClass { get; }
        object Parent { get; }
        NameSpace Session { get; }
        Attachments Attachments { get; }
        DateTime CreationTime { get; }
        DateTime LastModificationTime { get; }
        int Size { get; }
        string Categories { get; set; }
        OlImportance Importance { get; set; }
        OlSensitivity Sensitivity { get; set; }
        UserProperties UserProperties { get; }
        Actions Actions { get; }
        Links Links { get; }
        bool IsConflict { get; }
        OlDownloadState DownloadState { get; }
        OlMarkInterval MarkForDownload { get; set; }
        PropertyAccessor PropertyAccessor { get; }

        // Common Methods
        void Save();
        void SaveAs(string path, OlSaveAsType type = OlSaveAsType.olTXT);
        void Delete();
        void Display(object modal = null);
        IOutlookCommonItem Move(object destFolder);
        IOutlookCommonItem Copy();
        object GetInspector();
        void Close(OlInspectorClose SaveMode = OlInspectorClose.olSave);

        // Common Events
        event ItemEvents_10_OpenEventHandler Open;
        event ItemEvents_10_CloseEventHandler CloseEvent;
        event ItemEvents_10_ReadEventHandler Read;
        event ItemEvents_10_WriteEventHandler Write;
        event ItemEvents_10_BeforeDeleteEventHandler BeforeDelete;
        event ItemEvents_10_AttachmentAddEventHandler AttachmentAdd;
        event ItemEvents_10_AttachmentReadEventHandler AttachmentRead;
        event ItemEvents_10_AttachmentRemoveEventHandler AttachmentRemove;
        event ItemEvents_10_PropertyChangeEventHandler PropertyChange;
        event ItemEvents_10_CustomActionEventHandler CustomAction;
        event ItemEvents_10_CustomPropertyChangeEventHandler CustomPropertyChange;
        event ItemEvents_10_ForwardEventHandler Forward;
        event ItemEvents_10_ReplyEventHandler Reply;
        event ItemEvents_10_ReplyAllEventHandler ReplyAll;
    }
}