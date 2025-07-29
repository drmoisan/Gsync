using Newtonsoft.Json;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Gsync.Utilities.ReusableTypes;

namespace Gsync.Utilities.Interfaces
{
    public interface IScoDictionaryNew<TKey, TValue>: IConcurrentObservableDictionary<TKey, TValue>, ISmartSerializable<ScoDictionaryNew<TKey, TValue>>
    {
        void Notify([CallerMemberName] string propertyName = "");        
    }
}