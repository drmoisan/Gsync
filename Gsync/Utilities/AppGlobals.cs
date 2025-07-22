using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities
{
    public class AppGlobals
    {
        public AppGlobals(Application olApp) 
        { 
            OutlookApplication = olApp ?? throw new ArgumentNullException(nameof(olApp), "Outlook Application cannot be null.");
        }

        public Application OutlookApplication { get; internal set; }
    }
}
