using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using Gsync;

namespace Gsync.Utilities
{
    /// <summary>
    /// Tracks MailMessage objects and detects which have been deleted or moved (missing from the latest snapshot).
    /// Uses ImapOutlookItemWrapper.GetHashCode for consistent identity.
    /// </summary>
    public class MailMessageChangeTracker
    {
        private HashSet<int> _previousHashes;

        public MailMessageChangeTracker()
        {
            _previousHashes = new HashSet<int>();
        }

        /// <summary>
        /// Takes a snapshot of the current MailMessage list.
        /// </summary>
        public void Snapshot(IEnumerable<MailMessage> messages)
        {
            _previousHashes = new HashSet<int>(
                messages
                    .Where(m => m != null)
                    .Select(m => new ImapOutlookItemWrapper(m).GetHashCode())
            );
        }

        /// <summary>
        /// Returns the hash codes of messages that have been deleted or moved since the last snapshot.
        /// </summary>
        public List<int> GetDeletedOrMovedHashes(IEnumerable<MailMessage> currentMessages)
        {
            var currentHashes = new HashSet<int>(
                currentMessages
                    .Where(m => m != null)
                    .Select(m => new ImapOutlookItemWrapper(m).GetHashCode())
            );
            return _previousHashes.Where(hash => !currentHashes.Contains(hash)).ToList();
        }
    }
}
