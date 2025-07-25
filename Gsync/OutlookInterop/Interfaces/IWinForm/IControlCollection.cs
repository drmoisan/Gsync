using System.Collections;
using System.Windows.Forms;

namespace Gsync.OutlookInterop.Interfaces.IWinForm
{
    public interface IControlCollection
    {
        Control this[int index] { get; }
        Control this[string key] { get; }

        Control Owner { get; }

        void Add(Control value);
        void AddRange(Control[] controls);
        void Clear();
        bool Contains(Control control);
        bool ContainsKey(string key);
        Control[] Find(string key, bool searchAllChildren);
        int GetChildIndex(Control child);
        int GetChildIndex(Control child, bool throwException);
        IEnumerator GetEnumerator();
        int IndexOf(Control control);
        int IndexOfKey(string key);
        void Remove(Control value);
        void RemoveAt(int index);
        void RemoveByKey(string key);
        void SetChildIndex(Control child, int newIndex);
    }
}