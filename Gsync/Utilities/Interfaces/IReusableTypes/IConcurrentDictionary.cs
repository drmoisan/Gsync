using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.Interfaces
{
    public interface IConcurrentDictionary<TKey, TValue> : IDictionary<TKey, TValue>, ICollection<KeyValuePair<TKey, TValue>>, IEnumerable<KeyValuePair<TKey, TValue>>, IEnumerable
    {
        new TValue this[TKey key] { get; set; }
        new int Count { get; }
        bool IsEmpty { get; }
        new ICollection<TKey> Keys { get; }
        new ICollection<TValue> Values { get; }
        bool TryAdd(TKey key, TValue value);
        new bool ContainsKey(TKey key);
        bool TryRemove(TKey key, out TValue value);
        new bool TryGetValue(TKey key, out TValue value);
        bool TryUpdate(TKey key, TValue newValue, TValue comparisonValue);
        new void Clear();
        KeyValuePair<TKey, TValue>[] ToArray();
        TValue GetOrAdd(TKey key, Func<TKey, TValue> valueFactory);
        TValue GetOrAdd(TKey key, TValue value);
        TValue GetOrAdd<TArg>(TKey key, Func<TKey, TArg, TValue> valueFactory, TArg factoryArgument);
        TValue AddOrUpdate<TArg>(TKey key, Func<TKey, TArg, TValue> addValueFactory, Func<TKey, TValue, TArg, TValue> updateValueFactory, TArg factoryArgument);
        TValue AddOrUpdate(TKey key, Func<TKey, TValue> addValueFactory, Func<TKey, TValue, TValue> updateValueFactory);
        TValue AddOrUpdate(TKey key, TValue addValue, Func<TKey, TValue, TValue> updateValueFactory);
    }
}
