using System;
using System.ComponentModel;
using Gsync.Utilities.Interfaces;

namespace Gsync.Utilities.GTD
{
    public class ProjectElement : IProjectElement
    {
        private readonly object _syncRoot = new object();
        private string _id;
        private string _name;

        public string ID
        {
            get
            {
                lock (_syncRoot)
                {
                    return _id;
                }
            }
            set
            {
                bool changed = false;
                lock (_syncRoot)
                {
                    if (_id != value)
                    {
                        _id = value;
                        changed = true;
                    }
                }
                if (changed)
                    OnPropertyChanged(nameof(ID));
            }
        }

        public string Name
        {
            get
            {
                lock (_syncRoot)
                {
                    return _name;
                }
            }
            set
            {
                bool changed = false;
                lock (_syncRoot)
                {
                    if (_name != value)
                    {
                        _name = value;
                        changed = true;
                    }
                }
                if (changed)
                    OnPropertyChanged(nameof(Name));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            // Copy the delegate reference for thread safety
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }

        public int CompareTo(IProjectElement other)
        {
            if (other == null) return 1;
            string thisName, otherName;
            lock (_syncRoot)
            {
                thisName = _name;
            }
            otherName = other.Name; // Assume other is also thread-safe
            return string.Compare(thisName, otherName, StringComparison.Ordinal);
        }

        public bool Equals(IProjectElement other)
        {
            if (other == null) return false;
            string thisId, otherId;
            lock (_syncRoot)
            {
                thisId = _id;
            }
            otherId = other.ID; // Assume other is also thread-safe
            return string.Equals(thisId, otherId, StringComparison.Ordinal);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as IProjectElement);
        }

        public override int GetHashCode()
        {
            lock (_syncRoot)
            {
                return (_id != null ? _id.GetHashCode() : 0);
            }
        }
    }
}