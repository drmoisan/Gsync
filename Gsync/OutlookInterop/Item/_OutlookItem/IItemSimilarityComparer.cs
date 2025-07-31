using System;
using System.Collections.Generic;
using Gsync.OutlookInterop.Interfaces.Items;

namespace Gsync.OutlookInterop.Item
{
    public class IItemSimilarityComparer : IEqualityComparer<IItem>
    {
        public bool Equals(IItem x, IItem y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (x is null || y is null) return false;

            // IItem properties
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

            return true;
        }

        public int GetHashCode(IItem obj)
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
