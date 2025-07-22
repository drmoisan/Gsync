using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Gsync.Utilities.Extensions
{
    public static class JsonSerializerExtensions
    {
        public static JsonSerializer DeepCopy(this JsonSerializer serializer)
        {
            var copiedSerializer = new JsonSerializer
            {
                Context = serializer.Context,
                Culture = serializer.Culture,
                ContractResolver = serializer.ContractResolver,
                ConstructorHandling = serializer.ConstructorHandling,
                CheckAdditionalContent = serializer.CheckAdditionalContent,
                DateFormatHandling = serializer.DateFormatHandling,
                DateFormatString = serializer.DateFormatString,
                DateParseHandling = serializer.DateParseHandling,
                DateTimeZoneHandling = serializer.DateTimeZoneHandling,
                DefaultValueHandling = serializer.DefaultValueHandling,
                EqualityComparer = serializer.EqualityComparer,
                FloatFormatHandling = serializer.FloatFormatHandling,
                Formatting = serializer.Formatting,
                FloatParseHandling = serializer.FloatParseHandling,
                MaxDepth = serializer.MaxDepth,
                MetadataPropertyHandling = serializer.MetadataPropertyHandling,
                MissingMemberHandling = serializer.MissingMemberHandling,
                NullValueHandling = serializer.NullValueHandling,
                ObjectCreationHandling = serializer.ObjectCreationHandling,
                PreserveReferencesHandling = serializer.PreserveReferencesHandling,
                ReferenceResolver = serializer.ReferenceResolver,
                ReferenceLoopHandling = serializer.ReferenceLoopHandling,
                StringEscapeHandling = serializer.StringEscapeHandling,
                TraceWriter = serializer.TraceWriter,
                TypeNameHandling = serializer.TypeNameHandling,
                SerializationBinder = serializer.SerializationBinder,
                TypeNameAssemblyFormatHandling = serializer.TypeNameAssemblyFormatHandling
            };
            foreach (var converter in serializer.Converters)
            {
                copiedSerializer.Converters.Add(converter);
            }
            return copiedSerializer;
        }

        public static JsonSerializerSettings DeepCopy(this JsonSerializerSettings settings)
        {
            var copiedSerializer = new JsonSerializerSettings
            {
                Context = settings.Context,
                Culture = settings.Culture,
                ContractResolver = settings.ContractResolver,
                ConstructorHandling = settings.ConstructorHandling,
                CheckAdditionalContent = settings.CheckAdditionalContent,
                DateFormatHandling = settings.DateFormatHandling,
                DateFormatString = settings.DateFormatString,
                DateParseHandling = settings.DateParseHandling,
                DateTimeZoneHandling = settings.DateTimeZoneHandling,
                DefaultValueHandling = settings.DefaultValueHandling,
                EqualityComparer = settings.EqualityComparer,
                FloatFormatHandling = settings.FloatFormatHandling,
                Formatting = settings.Formatting,
                FloatParseHandling = settings.FloatParseHandling,
                MaxDepth = settings.MaxDepth,
                MetadataPropertyHandling = settings.MetadataPropertyHandling,
                MissingMemberHandling = settings.MissingMemberHandling,
                NullValueHandling = settings.NullValueHandling,
                ObjectCreationHandling = settings.ObjectCreationHandling,
                PreserveReferencesHandling = settings.PreserveReferencesHandling,
                ReferenceLoopHandling = settings.ReferenceLoopHandling,
                StringEscapeHandling = settings.StringEscapeHandling,
                TraceWriter = settings.TraceWriter,
                TypeNameHandling = settings.TypeNameHandling,
                SerializationBinder = settings.SerializationBinder,
                TypeNameAssemblyFormatHandling = settings.TypeNameAssemblyFormatHandling
            };
            foreach (var converter in settings.Converters)
            {
                copiedSerializer.Converters.Add(converter);
            }
            return copiedSerializer;
        }

        public static JsonSerializerSettings ExtractSettings(this JsonSerializer serializer)
        {
            if (serializer == null)
                throw new ArgumentNullException(nameof(serializer));

            var settings = new JsonSerializerSettings
            {
                Formatting = serializer.Formatting,
                DateFormatHandling = serializer.DateFormatHandling,
                DateTimeZoneHandling = serializer.DateTimeZoneHandling,
                DateParseHandling = serializer.DateParseHandling,
                FloatFormatHandling = serializer.FloatFormatHandling,
                FloatParseHandling = serializer.FloatParseHandling,
                StringEscapeHandling = serializer.StringEscapeHandling,
                Culture = serializer.Culture,
                MaxDepth = serializer.MaxDepth,
                CheckAdditionalContent = serializer.CheckAdditionalContent,
                ConstructorHandling = serializer.ConstructorHandling,
                ContractResolver = serializer.ContractResolver,
                DefaultValueHandling = serializer.DefaultValueHandling,
                MissingMemberHandling = serializer.MissingMemberHandling,
                NullValueHandling = serializer.NullValueHandling,
                ObjectCreationHandling = serializer.ObjectCreationHandling,
                PreserveReferencesHandling = serializer.PreserveReferencesHandling,
                ReferenceLoopHandling = serializer.ReferenceLoopHandling,
                TypeNameHandling = serializer.TypeNameHandling,
                MetadataPropertyHandling = serializer.MetadataPropertyHandling,
                TypeNameAssemblyFormatHandling = serializer.TypeNameAssemblyFormatHandling,
                Context = serializer.Context,
                SerializationBinder = serializer.SerializationBinder,
                TraceWriter = serializer.TraceWriter,
                EqualityComparer = serializer.EqualityComparer
            };

            return settings;
        }

    }
}
