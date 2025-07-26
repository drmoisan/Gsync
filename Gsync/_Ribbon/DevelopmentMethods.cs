using Gsync.Utilities.Interfaces;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Drawing.Design;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Ribbon
{
    internal class DevelopmentMethods
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #region ctor

        public DevelopmentMethods(IApplicationGlobals globals)
        {
            Globals = globals ?? throw new ArgumentNullException(nameof(globals), "Application Globals cannot be null.");
        }

        internal IApplicationGlobals Globals { get; set; }

        #endregion ctor

        #region Methods

        public void LoopInbox() 
        {
            var accountInboxes = Globals.StoresWrapper.Stores.Select(x => (x.Account, x.Inbox)).ToArray();
            foreach (var accountInbox in accountInboxes)
            {
                logger.Debug($"Processing Inbox: {accountInbox.Account.DisplayName} {accountInbox.Inbox.Name}");
                var items = accountInbox.Inbox.Items.Cast<dynamic>().Where(item => (item.MessageClass.ToString() as string).StartsWith("IPM.Schedule.Meeting.Resp")).ToArray();
                foreach (dynamic item in items)
                {
                    object itemObject = item;
                    var members = GetComTypeMembers(itemObject);
                    if (members.InteropType == null)
                    {
                        logger.Warn($"Unknown item type for MessageClass: {item.MessageClass}");
                    }
                    else 
                    { 
                        logger.Debug($"Item of message class {item.MessageClass} is of type {members.InteropType} and has the following\n" +
                            $"Properties: {string.Join(",", members.Properties)}\n" +
                            $"Methods:    {string.Join(",", members.Methods)}");                    
                    }
                    var meeting = item as MeetingItem;
                }
            }
        }
        
        public void MessageClassesInbox()
        {
            var accountInboxes = Globals.StoresWrapper.Stores.Select(x => (x.Account,x.Inbox)).ToArray();
            foreach (var accountInbox in accountInboxes)
            {
                logger.Debug($"Processing Inbox: {accountInbox.Account.DisplayName} {accountInbox.Inbox.Name}");
                var results = accountInbox.Inbox.Items.Cast<dynamic>().Select(item => item.MessageClass).Distinct().ToArray();
                foreach (dynamic result in results)
                {
                    logger.Debug($"Found Message Class: {result}");
                }
                                
            }

        }

        internal static Type GetKnownOutlookItemType(object comObject)
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

            return null;
        }
                
        public static (Type InteropType, List<string> Properties, List<string> Methods) GetComTypeMembers(object comObject)
        {
            var interopType = GetKnownOutlookItemType(comObject);
            if (interopType == null) { return (null, null, null); }

            var properties = interopType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                        .Select(p => p.Name)
                                        .Distinct()
                                        .ToList();

            var methods = interopType.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                                    .Where(m => !m.IsSpecialName)
                                    .Select(m => m.Name)
                                    .Distinct()
                                    .ToList();

            return (interopType, properties, methods);
        }

            
        

        #endregion Methods

    }
}
