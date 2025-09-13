using Microsoft.Office.Interop.Outlook;
using System;

namespace Gsync.OutlookInterop.Item
{
    public static class OutlookItemType
    {
        public static Type GetType(object comObject)
        {
            if (comObject == null) return null;

            if (comObject is AppointmentItem) return typeof(AppointmentItem);
            if (comObject is ContactItem) return typeof(ContactItem);
            if (comObject is DistListItem) return typeof(DistListItem);
            if (comObject is DocumentItem) return typeof(DocumentItem);
            if (comObject is JournalItem) return typeof(JournalItem);
            if (comObject is MailItem) return typeof(MailItem);
            if (comObject is MeetingItem) return typeof(MeetingItem);
            if (comObject is NoteItem) return typeof(NoteItem);
            if (comObject is PostItem) return typeof(PostItem);
            if (comObject is RemoteItem) return typeof(RemoteItem);
            if (comObject is ReportItem) return typeof(ReportItem);
            if (comObject is SharingItem) return typeof(SharingItem);
            if (comObject is StorageItem) return typeof(StorageItem);
            if (comObject is TaskItem) return typeof(TaskItem);
            if (comObject is TaskRequestAcceptItem) return typeof(TaskRequestAcceptItem);
            if (comObject is TaskRequestDeclineItem) return typeof(TaskRequestDeclineItem);
            if (comObject is TaskRequestItem) return typeof(TaskRequestItem);
            if (comObject is TaskRequestUpdateItem) return typeof(TaskRequestUpdateItem);

            return comObject.GetType();
        }
    }
}
