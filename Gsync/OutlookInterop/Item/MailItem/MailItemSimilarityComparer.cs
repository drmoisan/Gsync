using System;
using System.Collections.Generic;
using Gsync.OutlookInterop.Interfaces.Items;

namespace Gsync.OutlookInterop.Item
{
    public class MailItemSimilarityComparer : IEqualityComparer<IMailItem>
    {
        public bool Equals(IMailItem x, IMailItem y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (x is null || y is null) return false;

            // IItem properties excluding EntryID
            if (!StringEquals(x.BillingInformation, y.BillingInformation)) return false;
            if (!StringEquals(x.Body, y.Body)) return false;
            if (!StringEquals(x.Categories, y.Categories)) return false;
            if (x.Class != y.Class) return false;
            if (!StringEquals(x.Companies, y.Companies)) return false;
            if (!StringEquals(x.ConversationID, y.ConversationID)) return false;
            if (!DateEquals(x.CreationTime, y.CreationTime)) return false;            
            if (!StringEquals(x.HTMLBody, y.HTMLBody)) return false;
            if (x.Importance != y.Importance) return false;            
            if (!StringEquals(x.MessageClass, y.MessageClass)) return false;
            if (!StringEquals(x.Mileage, y.Mileage)) return false;
            if (x.NoAging != y.NoAging) return false;
            if (x.OutlookInternalVersion != y.OutlookInternalVersion) return false;
            if (!StringEquals(x.OutlookVersion, y.OutlookVersion)) return false;
            if (x.Saved != y.Saved) return false;
            if (!StringEquals(x.SenderEmailAddress, y.SenderEmailAddress)) return false;
            if (!StringEquals(x.SenderName, y.SenderName)) return false;
            if (x.Sensitivity != y.Sensitivity) return false;
            if (x.Size != y.Size) return false;
            if (!StringEquals(x.Subject, y.Subject)) return false;
            if (x.UnRead != y.UnRead) return false;

            // IMailItem properties
            if (!StringEquals(x.BCC, y.BCC)) return false;
            if (!StringEquals(x.CC, y.CC)) return false;
            if (!StringEquals(x.DeferredDeliveryTime, y.DeferredDeliveryTime)) return false;
            if (!StringEquals(x.DeleteAfterSubmit, y.DeleteAfterSubmit)) return false;
            if (!StringEquals(x.FlagRequest, y.FlagRequest)) return false;
            if (!StringEquals(x.ReceivedByName, y.ReceivedByName)) return false;
            if (!StringEquals(x.ReceivedOnBehalfOfName, y.ReceivedOnBehalfOfName)) return false;
            if (!DateEquals(x.ReceivedTime, y.ReceivedTime)) return false;
            if (!StringEquals(x.RecipientReassignmentProhibited, y.RecipientReassignmentProhibited)) return false;
            if (x.ReminderOverrideDefault != y.ReminderOverrideDefault) return false;
            if (x.ReminderPlaySound != y.ReminderPlaySound) return false;
            if (x.ReminderSet != y.ReminderSet) return false;
            if (!StringEquals(x.ReminderSoundFile, y.ReminderSoundFile)) return false;
            if (!DateEquals(x.ReminderTime, y.ReminderTime)) return false;
            if (!StringEquals(x.ReplyRecipientNames, y.ReplyRecipientNames)) return false;
            if (x.SaveSentMessageFolder != y.SaveSentMessageFolder) return false;
            if (!StringEquals(x.SenderEmailType, y.SenderEmailType)) return false;
            if (!StringEquals(x.SentOnBehalfOfName, y.SentOnBehalfOfName)) return false;
            if (!DateEquals(x.SentOn, y.SentOn)) return false;
            if (x.Submitted != y.Submitted) return false;
            if (!StringEquals(x.To, y.To)) return false;
            if (!StringEquals(x.VotingOptions, y.VotingOptions)) return false;
            if (!StringEquals(x.VotingResponse, y.VotingResponse)) return false;

            // Note: Recipients/ReplyRecipients, Attachments, etc. are intentionally omitted

            return true;
        }

        public int GetHashCode(IMailItem obj)
        {
            if (obj == null) return 0;

            // A simple way: combine all value fields in a hash
            int hash = 17;
            unchecked
            {
                hash = hash * 23 + StringHash(obj.BillingInformation);
                hash = hash * 23 + StringHash(obj.Body);
                hash = hash * 23 + StringHash(obj.Categories);
                hash = hash * 23 + obj.Class.GetHashCode();
                hash = hash * 23 + StringHash(obj.Companies);
                hash = hash * 23 + StringHash(obj.ConversationID);
                hash = hash * 23 + DateHash(obj.CreationTime);                
                hash = hash * 23 + StringHash(obj.HTMLBody);
                hash = hash * 23 + obj.Importance.GetHashCode();                
                hash = hash * 23 + StringHash(obj.MessageClass);
                hash = hash * 23 + StringHash(obj.Mileage);
                hash = hash * 23 + obj.NoAging.GetHashCode();
                hash = hash * 23 + obj.OutlookInternalVersion.GetHashCode();
                hash = hash * 23 + StringHash(obj.OutlookVersion);
                hash = hash * 23 + obj.Saved.GetHashCode();
                hash = hash * 23 + StringHash(obj.SenderEmailAddress);
                hash = hash * 23 + StringHash(obj.SenderName);
                hash = hash * 23 + obj.Sensitivity.GetHashCode();
                hash = hash * 23 + obj.Size.GetHashCode();
                hash = hash * 23 + StringHash(obj.Subject);
                hash = hash * 23 + obj.UnRead.GetHashCode();

                hash = hash * 23 + StringHash(obj.BCC);
                hash = hash * 23 + StringHash(obj.CC);
                hash = hash * 23 + StringHash(obj.DeferredDeliveryTime);
                hash = hash * 23 + StringHash(obj.DeleteAfterSubmit);
                hash = hash * 23 + StringHash(obj.FlagRequest);
                hash = hash * 23 + StringHash(obj.ReceivedByName);
                hash = hash * 23 + StringHash(obj.ReceivedOnBehalfOfName);
                hash = hash * 23 + DateHash(obj.ReceivedTime);
                hash = hash * 23 + StringHash(obj.RecipientReassignmentProhibited);
                hash = hash * 23 + obj.ReminderOverrideDefault.GetHashCode();
                hash = hash * 23 + obj.ReminderPlaySound.GetHashCode();
                hash = hash * 23 + obj.ReminderSet.GetHashCode();
                hash = hash * 23 + StringHash(obj.ReminderSoundFile);
                hash = hash * 23 + DateHash(obj.ReminderTime);
                hash = hash * 23 + StringHash(obj.ReplyRecipientNames);
                hash = hash * 23 + obj.SaveSentMessageFolder.GetHashCode();
                hash = hash * 23 + StringHash(obj.SenderEmailType);
                hash = hash * 23 + StringHash(obj.SentOnBehalfOfName);
                hash = hash * 23 + DateHash(obj.SentOn);
                hash = hash * 23 + obj.Submitted.GetHashCode();
                hash = hash * 23 + StringHash(obj.To);
                hash = hash * 23 + StringHash(obj.VotingOptions);
                hash = hash * 23 + StringHash(obj.VotingResponse);
            }
            return hash;
        }

        // Helpers for safe equality and hashing
        private static bool StringEquals(string a, string b) =>
            string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        private static int StringHash(string s) =>
            s == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(s);
        private static bool DateEquals(DateTime a, DateTime b) =>
            a == default && b == default ? true : a.Equals(b);
        private static int DateHash(DateTime dt) =>
            dt == default ? 0 : dt.GetHashCode();
    }
}
