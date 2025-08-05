
using Gsync.Utilities.ReusableTypes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Runtime;
using System.Text;

namespace Gsync.Utilities.HelperClasses.NewtonSoft
{
    public class ScoDictionaryConverter<TDerived, TKey, TValue> : JsonConverter<TDerived> where TDerived : ScoDictionaryNew<TKey, TValue>
    {
        public ScoDictionaryConverter() { }

        public override TDerived ReadJson(JsonReader reader, Type typeToConvert, TDerived existingValue, bool hasExistingValue, JsonSerializer serializer)
        {
            // First deserialize into JObject so we can control RemainingObject later
            var jObj = JObject.Load(reader);

            // Manually extract and deserialize CoDictionary and RemainingObject
            var coDict = jObj["CoDictionary"]?.ToObject<ConcurrentObservableDictionary<TKey, TValue>>(serializer);
            var remainingToken = jObj["RemainingObject"];

            var wrapper = new WrapperScoDictionary<TDerived, TKey, TValue>();
            wrapper.CoDictionary = coDict;

            // Dynamically compile the expected type
            Type dynamicType = wrapper.CompileType();

            // Deserialize RemainingObject into expected type
            wrapper.RemainingObject = remainingToken?.ToObject(dynamicType, serializer);

            // Final step
            return wrapper.ToDerived();

            //var wrapper = serializer.Deserialize(reader, typeof(WrapperScoDictionary<TDerived, TKey, TValue>)) as WrapperScoDictionary<TDerived, TKey, TValue>;
            //return wrapper?.ToDerived();            
        }

        public override void WriteJson(JsonWriter writer, TDerived value, JsonSerializer serializer)
        {
            var wrapper = new WrapperScoDictionary<TDerived, TKey, TValue>().ToComposition(value);
            serializer.Serialize(writer, wrapper);
        }

        public override bool CanRead => base.CanRead;

    }

    public class ScoDictionaryConverter : JsonConverter
    {
        public ScoDictionaryConverter() : base() { }

        public override bool CanConvert(Type objectType)
        {
            return objectType.IsDerivedFrom_ScoDictionaryNew();
            //return objectType.IsGenericType && (objectType.GetGenericTypeDefinition() == typeof(ScoDictionaryNew<,>) || objectType.GetGenericTypeDefinition() == typeof(WrapperScoDictionary<,,>));
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            // Create the generic type WrapperScoDictionary<TDerived, TKey, TValue> at runtime
            Type[] genericArguments = objectType.GetScoDictionaryNewGenerics();
            Type wrapperType = typeof(WrapperScoDictionary<,,>).MakeGenericType(objectType, genericArguments[0], genericArguments[1]);
            var wrapper = serializer.Deserialize(reader, wrapperType);
            return wrapperType.GetMethod("ToDerived", [])?.Invoke(wrapper, null);
            
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            Type valueType = value.GetType();
            Type[] genericArguments = valueType.GetScoDictionaryNewGenerics();
            //Type[] genericArguments = valueType.GetGenericArguments();
            Type wrapperType = typeof(WrapperScoDictionary<,,>).MakeGenericType(valueType, genericArguments[0], genericArguments[1]);
            var wrapper = Activator.CreateInstance(wrapperType);
            var toComposition = wrapperType.GetMethod("ToComposition");
            toComposition?.Invoke(wrapper, [value]);
            serializer.Serialize(writer, wrapper, wrapperType);

            //var wrapper2 = new WrapperScoDictionary<ScoDictionaryNew<string,string>,string,string> ();
                        
        }
    }
}
