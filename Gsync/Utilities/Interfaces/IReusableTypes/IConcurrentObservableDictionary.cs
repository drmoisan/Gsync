using System;
using System.Collections.Generic;
using Gsync.Utilities.ReusableTypes;
using Gsync.Utilities.Interfaces;

namespace Gsync.Utilities.Interfaces
{
    public interface IConcurrentObservableDictionary<TKey, TValue>:IConcurrentDictionary<TKey, TValue>
    {
        event EventHandler<DictionaryChangedEventArgs<TKey, TValue>> CollectionChanged;

        new TValue AddOrUpdate(TKey key, Func<TKey, TValue> addValueFactory, Func<TKey, TValue, TValue> updateValueFactory);
        TValue AddOrUpdate(TKey key, TValue value);
        new TValue AddOrUpdate(TKey key, TValue addValue, Func<TKey, TValue, TValue> updateValueFactory);
        Dictionary<TKey, HashSet<IDictionaryObserver<TKey, TValue>>> AddPartialObserver(Action<DictionaryChangedEventArgs<TKey, TValue>> action, params TKey[] keys);
        Dictionary<TKey, HashSet<IDictionaryObserver<TKey, TValue>>> AddPartialObserver(IDictionaryObserver<TKey, TValue> observer, params TKey[] keys);
        new void Clear();
        new TValue GetOrAdd(TKey key, Func<TKey, TValue> valueFactory);
        new TValue GetOrAdd(TKey key, TValue value);
        Dictionary<TKey, HashSet<IDictionaryObserver<TKey, TValue>>> RemoveAllObservers();
        Dictionary<TKey, HashSet<IDictionaryObserver<TKey, TValue>>> RemovePartialObserver(IDictionaryObserver<TKey, TValue> observer);
        Dictionary<TKey, HashSet<IDictionaryObserver<TKey, TValue>>> RemovePartialObserver(IDictionaryObserver<TKey, TValue> observer, params TKey[] keys);
        Dictionary<TKey, HashSet<IDictionaryObserver<TKey, TValue>>> RemovePartialObserver(params TKey[] keys);
        new bool TryAdd(TKey key, TValue value);
        new bool TryRemove(TKey key, out TValue value);
        new bool TryUpdate(TKey key, TValue newValue, TValue comparisonValue);
    }
}