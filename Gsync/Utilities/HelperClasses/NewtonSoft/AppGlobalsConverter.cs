using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using Gsync.Utilities.Interfaces;
using Gsync.Utilities.ReusableTypes;

namespace Gsync.Utilities.HelperClasses.NewtonSoft
{
    public class AppGlobalsConverter: JsonConverter<IApplicationGlobals>
    {
        private IApplicationGlobals _globals;

        public AppGlobalsConverter(IApplicationGlobals globals)
        {
            _globals = globals;
        }

        public override IApplicationGlobals ReadJson(JsonReader reader, Type objectType, IApplicationGlobals existingValue, bool hasExistingValue, JsonSerializer serializer)
        {
            return _globals;
        }

        public override void WriteJson(JsonWriter writer, IApplicationGlobals value, JsonSerializer serializer)
        {
            object message = "default";
            serializer.Serialize(writer, message);
        }
    }
}
