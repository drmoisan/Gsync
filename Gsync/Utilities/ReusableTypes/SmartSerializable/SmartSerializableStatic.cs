using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Gsync.Utilities.Interfaces;


namespace Gsync.Utilities.ReusableTypes
{
    
    public static class SmartSerializableStatic
    {
        public static bool IsSmartSerializable(this Type type)
        {
            return type.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ISmartSerializable<>));
        }
    }

    #region Old Static Class

    //public static class SmartSerializable
    //{
    //    private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
    //        System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

    //    public static bool IsSmartSerializable<T>(T instance)
    //    {
    //        return instance.GetType().IsSmartSerializable();
    //    }

    //    public static bool IsSmartSerializable(this Type type)
    //    {
    //        return type.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ISmartSerializable<>));
    //    }

    //    private static SmartSerializableBase GetInstance() => new();

    //    private static SmartSerializable<T> GetInstance<T>()
    //        where T : class, ISmartSerializable<T>, new() => new();

    //    public static T Deserialize<T>(string fileName, string folderPath)
    //        where T : class, ISmartSerializable<T>, new() =>
    //        GetInstance<T>().Deserialize(fileName, folderPath);

    //    public static T Deserialize<T>(string fileName, string folderPath, bool askUserOnError)
    //        where T : class, ISmartSerializable<T>, new() =>
    //        GetInstance<T>().Deserialize(fileName, folderPath, askUserOnError);

    //    public static T Deserialize<T>(string fileName, string folderPath, bool askUserOnError, JsonSerializerSettings settings)
    //        where T : class, ISmartSerializable<T>, new() =>
    //        GetInstance<T>().Deserialize(fileName, folderPath, askUserOnError, settings);

    //    public static T Deserialize<T, U>(SmartSerializable<U> config)
    //        where T : class, ISmartSerializable<T>, new()
    //        where U : class, ISmartSerializable<U>, new() =>
    //        GetInstance<T>().Deserialize(config);

    //    public static T DeserializeObject<T>(string json, JsonSerializerSettings settings) where T : class
    //    {
    //        T instance = default;
    //        try
    //        {
    //            instance = JsonConvert.DeserializeObject<T>(json, settings);
    //        }
    //        catch (Exception e)
    //        {
    //            logger.Error(e.Message, e);
    //        }
    //        if (instance is not null && IsSmartSerializable(instance))
    //        {
    //            var config = typeof(T).GetProperty("Config").GetValue(instance) as NewSmartSerializableConfig;
    //            if (config is not null)
    //            {
    //                config.JsonSettings = settings.DeepCopy();
    //            }
    //        }

    //        return instance;
    //    }

    //    //public static async Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config)
    //    //    where T : class, ISmartSerializable<T>, new()
    //    //    where U : class, ISmartSerializable<U>, new() =>
    //    //    await GetInstance<T>().DeserializeAsync(config);

    //    public static async Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config)
    //        where T : class, new()
    //        where U : class, ISmartSerializable<U>, new() =>
    //        await GetInstance().DeserializeAsync<T,U>(config);

    //    public static async Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config, bool askUserOnError)
    //        where T : class, new()
    //        where U : class, ISmartSerializable<U>, new() =>
    //        await GetInstance().DeserializeAsync<T,U>(config, askUserOnError);

    //    public static async Task<T> DeserializeAsync<T, U>(SmartSerializable<U> config, bool askUserOnError, Func<T> altLoader)
    //        where T : class, new()
    //        where U : class, ISmartSerializable<U>, new() =>
    //        await GetInstance().DeserializeAsync<T, U>(config, askUserOnError, altLoader);

    //}

    #endregion Old Static Class
}
