using System;
using System.Net.Mail;
using System.Linq;

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
        /// </summary>
        /// <param name="mailMessage">The MailMessage to wrap.</param>
        /// <param name="imapUid">Optional IMAP UID if available.</param>
        public ImapOutlookItemWrapper(MailMessage mailMessage, string imapUid = null)
        {
            if (mailMessage == null)
                throw new ArgumentNullException(nameof(mailMessage));

            // Message-ID is not directly exposed by MailMessage, but may be in Headers
            MessageId = mailMessage.Headers?["Message-ID"];
            Subject = mailMessage.Subject;
            From = mailMessage.From?.Address;
            To = string.Join(";", mailMessage.To.Select(addr => addr.Address));
            Date = ParseDateHeader(mailMessage.Headers?["Date"]);
            ImapUid = imapUid;
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