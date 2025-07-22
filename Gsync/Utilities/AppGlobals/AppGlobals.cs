using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Gsync.Utilities.Interfaces;

namespace Gsync.Utilities
{
    public class AppGlobals : IApplicationGlobals
    {
        #region ctor

        public AppGlobals(Application olApp)
        {
            OutlookApplication = olApp ?? throw new ArgumentNullException(nameof(olApp), "Outlook Application cannot be null.");
        }

        public async Task<AppGlobals> InitAsync()
        {
            await Task.CompletedTask; // Simulate async initialization 
            return this;
        }

        public static async Task<AppGlobals> CreateAsync(Application olApp)
        {
            if (olApp == null)
                throw new ArgumentNullException(nameof(olApp), "Outlook Application cannot be null.");
            var globals = new AppGlobals(olApp);
            return await globals.InitAsync();
        }

        #endregion ctor

        #region Public Properties

        public Application OutlookApplication { get; internal set; }

        public IFileSystemFolderPaths FS { get; internal set; }

        #endregion Public Properties

    }
}
