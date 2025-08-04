using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook; 
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Gsync.Utilities.ReusableTypes;
using Gsync.Utilities.Interfaces;
using System.Runtime.Serialization;
using System.Threading;

namespace Gsync.OutlookInterop
{
    public class StoresWrapper: SmartSerializable<StoresWrapper>
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #region ctor

        public StoresWrapper(): base() { base._parent = this; }

        public StoresWrapper(IApplicationGlobals globals)
        {
            base._parent = this;
            Globals = globals;
        }

        public virtual StoresWrapper Init()
        {
            var olApp = Globals.OutlookApplication;

            Stores = olApp.Session.Stores
                .Cast<Outlook.Store>()
                .Select(store => new StoreWrapper(store)).ToList();

            Stores.ForEach(store => logger.Debug($"Store found: {store.DisplayName}, {store.InnerStore.StoreID}"));

            foreach (Outlook.Account account in olApp.Session.Accounts)
            {
                logger.Debug($"Processing account {account.DisplayName} with delivery store ID: {account.DeliveryStore.StoreID}");
                var match = Stores.Find(store => store.InnerStore.StoreID == account.DeliveryStore.StoreID);
                if (match is not null)
                {
                    logger.Debug($"Matched store {match.DisplayName} to account {account.DisplayName}");
                    match.Account = account;
                }
                else { logger.Debug($"No match found for account {account.DisplayName}"); }
            }

            Stores.ForEach(store => store.Init());
            
            //var deliveryStores = olApp.Session.Accounts
            //    .Cast<Outlook.Account>()
            //    .Where(account => account.AccountType == OlAccountType.olImap)
            //    .Select(account => new StoreWrapper(account.DeliveryStore, account).Init())
            //    .ToList();

            return this;
        }

        public static async Task<StoresWrapper> CreateAsync(IApplicationGlobals globals, CancellationToken cancel)
        {
            return await Task.Run(() => new StoresWrapper(globals).Init(), cancel);
        }

        [OnDeserialized]
        public async void RewireOlObjects(System.Runtime.Serialization.StreamingContext context)
        {
            try
            {
                await RewireOlObjectsAsync(context);
            }
            catch (System.Exception e)
            {
                logger.Error($"Error in {nameof(RewireOlObjects)}: {e.Message}");                
            }
        }

        internal async Task RewireOlObjectsAsync(StreamingContext context)
        {
            this.Stores ??= [];           
            //var stores = Globals.Ol.NamespaceMAPI.Stores
            var stores = Globals.OutlookApplication.GetNamespace("MAPI")
                .Stores
                .Cast<Outlook.Store>()                
                .Where(store => store.ExchangeStoreType != OlExchangeStoreType.olExchangePublicFolder);

            foreach (var store in stores)
            {
                
                var storeWrapper = Stores.Find(x => x.DisplayName == store.DisplayName);
                if (storeWrapper is null)
                {
                    storeWrapper = await Task.Run(() => new StoreWrapper(store).Init());
                    Stores.Add(storeWrapper);
                }
                else
                {
                    await Task.Run(() => storeWrapper.Restore(store));
                    //await Task.Run(() => storeWrapper.RestoreGlobalAddresses(Globals.Ol.App));
                    
                }                                
            }
        }

        #endregion ctor

        [JsonProperty]
        internal IApplicationGlobals Globals { get; set; }

        [JsonProperty]
        public List<StoreWrapper> Stores { get; set; } 



    }
}
