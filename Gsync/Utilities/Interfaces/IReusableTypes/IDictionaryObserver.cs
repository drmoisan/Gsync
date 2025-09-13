using Gsync.Utilities.ReusableTypes;

namespace Gsync.Utilities.Interfaces
{
    public interface IDictionaryObserver<TKey, TValue>
    {
        void OnEventOccur(DictionaryChangedEventArgs<TKey, TValue> args);
    }
}