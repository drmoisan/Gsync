using Gsync.Utilities.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
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
            var inboxes = Globals.StoresWrapper.Stores.Select(x => x.Inbox).ToArray();
            foreach (var inbox in inboxes)
            {
                logger.Debug($"Processing Inbox: {inbox.Name}");
                foreach (var item in inbox.Items)
                {
                    if (item is Microsoft.Office.Interop.Outlook.MailItem mailItem)
                    {
                        logger.Debug($"Found Mail Item: {mailItem.Subject}");
                        // Process the mail item as needed
                    }
                }
            }

            #endregion Methods

        }
    }
}
