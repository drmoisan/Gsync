using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using System.Threading;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using Gsync.Utilities.Extensions;
using Gsync.Utilities.HelperClasses.NewtonSoft;

namespace Gsync.Utilities.ReusableTypes
{
    public class SmartSerializableLoader : SmartSerializable<SmartSerializableLoader>
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
        System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public SmartSerializableLoader() : base() { base._parent = this; }
        public SmartSerializableLoader(AppGlobals globals)
        {
            base._parent = this;
            Globals = globals;
            ResetLazy();
        }

        private void ResetLazy()
        {
            base.Config.ResetLazy(
                localJsonSettings: new Lazy<JsonSerializerSettings>(GetSettings),
                netJsonSettings: new Lazy<JsonSerializerSettings>(GetSettings),
                jsonSettings: new Lazy<JsonSerializerSettings>(GetSettings));
        }

        protected bool _engine;
        public bool Engine
        {
            get => _engine;
            set { _engine = value; Notify(); }
        }

        [JsonProperty]
        internal AppGlobals Globals { get => _globals; set => _globals = value; }
        private AppGlobals _globals;

        public Type T { get => _t; set { _t = value; Notify(); } }
        private Type _t;

        private JsonSerializerSettings GetSettings()
        {
            Globals.ThrowIfNull();
            var settings = GetDefaultSettings();
            settings.PreserveReferencesHandling = PreserveReferencesHandling.All;
            settings.Converters.Add(new AppGlobalsConverter(Globals));
            settings.Converters.Add(new FilePathHelperConverter(Globals.FS));
            return settings;
        }

        public static async Task<SmartSerializableLoader> DeserializeAsync(
            AppGlobals globals, string jsonObject, CancellationToken cancel = default)
        {
            try
            {
                if (globals is null) { throw new ArgumentNullException(nameof(globals)); }
                var loader = new SmartSerializableLoader(globals);
                return await Task.Run(() => loader.DeserializeConfig(jsonObject), cancel);
            }
            catch (ArgumentNullException e)
            {
                logger.Error($"Error in {nameof(DeserializeAsync)}. {nameof(globals)} cannot be null\n{e.Message}", e);
                throw;
            }

            catch (TaskCanceledException)
            {
                logger.Warn("Task was cancelled.");
                return null;
            }
            catch (Exception e)
            {
                logger.Error($"Error in {nameof(DeserializeAsync)}.\n{e.Message}", e);
                throw;
            }
        }

        public static async Task<SmartSerializableLoader> DeserializeAsync(
            AppGlobals globals, byte[] binary, CancellationToken cancel = default)
        {                
            try
            {
                if (globals is null) { throw new ArgumentNullException(nameof(globals)); }
                var loader = new SmartSerializableLoader(globals);
                return await Task.Run(() => loader.DeserializeConfig(binary), cancel);
            }
            catch (ArgumentNullException e)
            {
                logger.Error($"Error in {nameof(DeserializeAsync)}. {nameof(globals)} cannot be null\n{e.Message}", e);
                throw;
            }

            catch (TaskCanceledException)
            {
                logger.Warn("Task was cancelled.");
                return null;
            }
            catch (Exception e)
            {
                logger.Error($"Error in {nameof(DeserializeAsync)}.\n{e.Message}", e);
                throw;
            }                
        }
        
        internal SmartSerializableLoader DeserializeConfig(byte[] binary)
        {
            var jsonObject = TryConvertBinaryToJson(binary);
            if (jsonObject.IsNullOrEmpty())
            {
                return null;
            }
            else
            {
                return DeserializeConfig(jsonObject);
            }
        }

        private SmartSerializableLoader DeserializeConfig(string jsonObject)
        {
            var settings = GetSettings();
            SmartSerializableLoader instance = null;
            try
            {
                instance = JsonConvert.DeserializeObject<SmartSerializableLoader>(
                    jsonObject, settings);
            }
            catch (Exception e)
            {
                logger.Error($"Error in {nameof(DeserializeConfig)}.\n{e.Message}", e);
                return null;
            }
                
            instance.Globals = Globals;
            instance.ResetLazy();
            instance.Config.ActivateMostRecent();
            return instance;
        }

        internal string TryConvertBinaryToJson(byte[] binary)
        {
            try
            {
                var jsonObject = System.Text.Encoding.UTF8.GetString(binary);
                return jsonObject;
            }
            catch (Exception e)
            {
                logger.Error($"Error in {nameof(TryConvertBinaryToJson)}.\n{e.Message}",e);
                return null;
            }
        }

        
    }
}


