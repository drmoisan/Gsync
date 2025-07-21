using System;
using System.Linq;
using System.Net.Mail;
using Microsoft.Office.Interop.Outlook;

namespace Gsync
{
    /// <summary>
    /// Wraps an Outlook item using only IMAP-accessible properties.
    /// </summary>
    public class ImapOutlookItemWrapper : IEquatable<ImapOutlookItemWrapper>
    {
        public string MessageId { get; }
        public string Subject { get; }
        public string From { get; }
        public string To { get; }
        public DateTimeOffset Date { get; }
        public string ImapUid { get; }

        public ImapOutlookItemWrapper(string messageId, string subject, string from, string to, DateTimeOffset date, string imapUid)
        {
            MessageId = messageId;
            Subject = subject;
            From = from;
            To = to;
            Date = date;
            ImapUid = imapUid;
        }

        /// <summary>
        /// Constructs an ImapOutlookItemWrapper from a MailMessage.
        /// Attempts to extract the IMAP UID from the MailMessage headers if present.
        /// </summary>
        /// <param name="mailMessage">The MailMessage to wrap.</param>
        public ImapOutlookItemWrapper(MailMessage mailMessage)
        {
            if (mailMessage == null)
                throw new ArgumentNullException(nameof(mailMessage));

            MessageId = mailMessage.Headers?["Message-ID"];
            Subject = mailMessage.Subject;
            From = mailMessage.From?.Address;
            To = string.Join(";", mailMessage.To.Select(addr => addr.Address));
            Date = ParseDateHeader(mailMessage.Headers?["Date"]);
            ImapUid = ExtractImapUidFromMailMessage(mailMessage);
        }

        /// <summary>
        /// Constructs an ImapOutlookItemWrapper from a Microsoft.Office.Interop.Outlook.MailItem.
        /// Attempts to extract the IMAP UID from the MailItem's internet headers if present.
        /// </summary>
        /// <param name="mailItem">The Outlook MailItem to wrap.</param>
        public ImapOutlookItemWrapper(MailItem mailItem)            
        {
            if (mailItem == null)
                throw new ArgumentNullException(nameof(mailItem));

            MessageId = GetMessageIdFromMailItem(mailItem);
            Subject = mailItem.Subject;
            From = GetSenderEmailAddress(mailItem);
            To = mailItem.Recipients != null
                ? string.Join(";", mailItem.Recipients
                    .Cast<Recipient>()
                    .Where(r => r != null && !string.IsNullOrEmpty(r.Address))
                    .Select(r => r.Address))
                : string.Empty;
            Date = mailItem.SentOn != null && mailItem.SentOn != DateTime.MinValue
                ? new DateTimeOffset(mailItem.SentOn)
                : DateTimeOffset.MinValue;
            ImapUid = ExtractImapUidFromMailItem(mailItem);
        }

        // Internal for testability
        internal ImapOutlookItemWrapper(MailItem mailItem, Func<MailItem, string> senderExtractor)
        {
            if (mailItem == null)
                throw new ArgumentNullException(nameof(mailItem));

            MessageId = GetMessageIdFromMailItem(mailItem);
            Subject = mailItem.Subject;
            From = (senderExtractor ?? GetSenderEmailAddress)(mailItem);
            To = mailItem.Recipients != null
                ? string.Join(";", mailItem.Recipients
                    .Cast<Recipient>()
                    .Where(r => r != null && !string.IsNullOrEmpty(r.Address))
                    .Select(r => r.Address))
                : string.Empty;
            Date = mailItem.SentOn != null && mailItem.SentOn != DateTime.MinValue
                ? new DateTimeOffset(mailItem.SentOn)
                : DateTimeOffset.MinValue;
            ImapUid = ExtractImapUidFromMailItem(mailItem);
        }

        // Default sender extraction logic
        protected virtual string GetSenderEmailAddress(MailItem mailItem)
        {
            return mailItem.SenderEmailAddress;
        }

        internal string ExtractImapUidFromMailMessage(MailMessage mailMessage)
        {
            // Common custom header for IMAP UID is "X-IMAP-UID" or "Imap-Uid"
            var uid = mailMessage.Headers?["X-IMAP-UID"] ?? mailMessage.Headers?["Imap-Uid"];
            return string.IsNullOrEmpty(uid) ? null : uid;
        }

        internal string ExtractImapUidFromMailItem(MailItem mailItem)
        {
            // Try to extract IMAP UID from the internet headers if present
            const string PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
            try
            {
                var headers = mailItem.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS) as string;
                if (!string.IsNullOrEmpty(headers))
                {
                    foreach (var line in headers.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
                    {
                        if (line.StartsWith("X-IMAP-UID:", StringComparison.OrdinalIgnoreCase))
                        {
                            return line.Substring("X-IMAP-UID:".Length).Trim();
                        }
                        if (line.StartsWith("Imap-Uid:", StringComparison.OrdinalIgnoreCase))
                        {
                            return line.Substring("Imap-Uid:".Length).Trim();
                        }
                    }
                }
            }
            catch
            {
                // Property may not exist or be accessible
            }
            return null;
        }

        private static string GetMessageIdFromMailItem(MailItem mailItem)
        {
            // PR_TRANSPORT_MESSAGE_HEADERS property tag
            const string PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
            try
            {
                var headers = mailItem.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS) as string;
                if (!string.IsNullOrEmpty(headers))
                {
                    foreach (var line in headers.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
                    {
                        if (line.StartsWith("Message-ID:", StringComparison.OrdinalIgnoreCase))
                        {
                            return line.Substring("Message-ID:".Length).Trim();
                        }
                    }
                }
            }
            catch
            {
                // Property may not exist or be accessible
            }
            return null;
        }

        private static DateTimeOffset ParseDateHeader(string dateHeader)
        {
            if (DateTimeOffset.TryParse(dateHeader, out var result))
                return result;
            return DateTimeOffset.MinValue;
        }

        #region IEquatable<ImapOutlookItemWrapper> Members

        public bool Equals(ImapOutlookItemWrapper other)
        {
            if (other == null)
                return false;

            // Message-ID is globally unique for emails; fallback to other properties if missing
            if (!string.IsNullOrEmpty(MessageId) && !string.IsNullOrEmpty(other.MessageId))
                return string.Equals(MessageId, other.MessageId, StringComparison.OrdinalIgnoreCase);

            // Fallback: compare other IMAP properties
            return string.Equals(Subject, other.Subject, StringComparison.OrdinalIgnoreCase)
                && string.Equals(From, other.From, StringComparison.OrdinalIgnoreCase)
                && string.Equals(To, other.To, StringComparison.OrdinalIgnoreCase)
                && Date.Equals(other.Date);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as ImapOutlookItemWrapper);
        }

        public override int GetHashCode()
        {
            // Prefer Message-ID for hash code if available
            if (!string.IsNullOrEmpty(MessageId))
                return MessageId.ToLowerInvariant().GetHashCode();

            // Fallback: combine other properties using a simple hash code algorithm
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + (Subject?.ToLowerInvariant().GetHashCode() ?? 0);
                hash = hash * 23 + (From?.ToLowerInvariant().GetHashCode() ?? 0);
                hash = hash * 23 + (To?.ToLowerInvariant().GetHashCode() ?? 0);
                hash = hash * 23 + Date.GetHashCode();
                return hash;
            }
        }

        #endregion IEquatable<ImapOutlookItemWrapper> Members
    }
}