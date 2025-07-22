using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Extensions
{
    public static class LazyExtension
    {
        public static Lazy<T> ToLazy<T>(this T value) where T: class
        {
            return new Lazy<T>(value.Return);
        }

        public static Lazy<T> ToLazyValue<T>(this T value) where T : struct
        {
            return new Lazy<T>(() => value);
        }

        //public static LazyTry<T> ToLazyTry<T>(this T value) where T : class
        //{
        //    return new LazyTry<T>(value.Return);
        //}

        //public static LazyTry<T> ToLazyTryValue<T>(this T value) where T : struct
        //{
        //    return new LazyTry<T>(() => value);
        //}

        public static Func<T> AsFunc<T>(this T value) where T: class
        {
            return new Func<T>(value.Return);
        }        

        internal static T Return<T>(this T value) where T : class
        {
            return value;
        }

        //public static AsyncLazy<T> ToAsyncLazy<T>(this T value) where T : class
        //{
        //    return new AsyncLazy<T>(() => value);
        //}
    }
}
