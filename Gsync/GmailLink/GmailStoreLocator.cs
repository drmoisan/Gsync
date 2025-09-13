using System;
using System.Collections.Generic;
using System.Linq;
using Gsync.OutlookInterop;

namespace Gsync.GmailLink
{
    public class GmailStoreLocator
    {
        private readonly StoresWrapper _storesWrapper;

        public GmailStoreLocator(StoresWrapper storesWrapper)
        {
            _storesWrapper = storesWrapper ?? throw new ArgumentNullException(nameof(storesWrapper));
        }

        /// <summary>
        /// Returns all stores that appear to be Gmail accounts based on email address and metadata.
        /// </summary>
        public List<StoreWrapper> GetGmailStores()
        {
            return _storesWrapper.Stores
                .Where(IsGmailStore)
                .ToList();
        }

        /// <summary>
        /// Heuristic to determine if a store is Gmail-backed.
        /// </summary>
        private bool IsGmailStore(StoreWrapper store)
        {
            if (store == null) return false;

            var email = store.UserEmailAddress?.ToLowerInvariant() ?? string.Empty;
            var display = store.DisplayName?.ToLowerInvariant() ?? string.Empty;

            bool looksLikeGmail = email.Contains("@gmail.com") || email.Contains("@googlemail.com");
            bool namedLikeGmail = display.Contains("gmail") || display.Contains("google");

            bool accountIsImap = store.Account?.AccountType == Microsoft.Office.Interop.Outlook.OlAccountType.olImap;

            return looksLikeGmail || (accountIsImap && namedLikeGmail);
        }
    }
}
