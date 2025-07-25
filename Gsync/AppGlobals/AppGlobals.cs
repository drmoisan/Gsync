using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Gsync.Utilities.Interfaces;
using System.Threading;
using Gsync.Utilities.Threading;
using Gsync.OutlookInterop;

namespace Gsync
{
    public class AppGlobals : IApplicationGlobals
    {
        #region ctor

        public AppGlobals(Application olApp)
        {
            OutlookApplication = olApp ?? throw new ArgumentNullException(nameof(olApp), "Outlook Application cannot be null.");
        }

        public async Task<AppGlobals> InitAsync(SynchronizationContext context, int uiThreadId)
        {
            if (SynchronizationContext.Current is null) { SynchronizationContext.SetSynchronizationContext(context); }

            await Task.Run(() => UI = new UiWrapper(context, uiThreadId)).ConfigureAwait(true);
            await Task.Run(() => FS = new AppFileSystemFolderPaths()).ConfigureAwait(true);            
            this.StoresWrapper = new StoresWrapper(this).Init();

            return this;
        }

        public static async Task<AppGlobals> CreateAsync(Application olApp, SynchronizationContext context, int uiThreadId)
        {
            if (olApp == null)
                throw new ArgumentNullException(nameof(olApp), "Outlook Application cannot be null.");
            var globals = new AppGlobals(olApp);            
            return await globals.InitAsync(context, uiThreadId);
        }

        #endregion ctor

        #region Public Properties

        public Application OutlookApplication { get; internal set; }

        public IFileSystemFolderPaths FS { get; internal set; }

        public UiWrapper UI { get; internal set; }    
        
        public StoresWrapper StoresWrapper { get; internal set; }

        #endregion Public Properties

    }
}
