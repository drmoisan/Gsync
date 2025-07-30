using Gsync.OutlookInterop.Item;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Test.OutlookInterop.Item
{
    public class TestableOutlookItemWrapper : OutlookItemWrapper
    {
        public List<object> ReleasedObjects = new List<object>();        
        public bool ThrowOnRelease = false;
        public bool IsComObject = true;

        public TestableOutlookItemWrapper(object item) : base(item) { }

        public TestableOutlookItemWrapper(object item, ItemEvents_10_Event events = null)
            : base(item, events) { Init(); }

        public TestableOutlookItemWrapper(object item, ItemEvents_10_Event events, ImmutableHashSet<string> supportedTypes)
            : base(item, events, supportedTypes) { Init(); }

        public void InvokeOnAttachmentAdd(Attachment a)
        {
            var bt = base.GetType().BaseType;
            var mi = bt.GetMethod("OnAttachmentAdd", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            mi.Invoke(this, new object[] { a });
        }
                
        //public void InvokeOnAttachmentAdd(Attachment a) => base.GetType()
        //        .GetMethod("OnAttachmentAdd", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
        //        .Invoke(this, new object[] { a });

        public void InvokeOnAttachmentRead(Attachment a) => base.GetType().BaseType
            .GetMethod("OnAttachmentRead", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
            .Invoke(this, new object[] { a });

        public void InvokeOnAttachmentRemove(Attachment a) => base.GetType().BaseType
            .GetMethod("OnAttachmentRemove", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
            .Invoke(this, new object[] { a });

        public void InvokeOnBeforeDelete(object item, ref bool cancel)
        {
            var args = new object[] { item, cancel };
            base.GetType().BaseType
                .GetMethod("OnBeforeDelete", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Invoke(this, args);
            cancel = (bool)args[1];
        }

        public void InvokeOnCloseEvent(ref bool cancel)
        {
            var args = new object[] { cancel };
            base.GetType().BaseType
                .GetMethod("OnCloseEvent", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Invoke(this, args);
            cancel = (bool)args[0];
        }

        public void InvokeOnOpen(ref bool cancel)
        {
            var args = new object[] { cancel };
            base.GetType().BaseType
                .GetMethod("OnOpen", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Invoke(this, args);
            cancel = (bool)args[0];
        }

        public void InvokeOnPropertyChange(string name) => base.GetType().BaseType
            .GetMethod("OnPropertyChange", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
            .Invoke(this, new object[] { name });

        public void InvokeOnRead() => base.GetType().BaseType
            .GetMethod("OnRead", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
            .Invoke(this, null);

        public void InvokeOnWrite(ref bool cancel)
        {
            var args = new object[] { cancel };
            base.GetType().BaseType
                .GetMethod("OnWrite", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Invoke(this, args);
            cancel = (bool)args[0];
        }

        protected override void ReleaseComObject(object comObj)
        {
            if (comObj != null && IsComObjectFunc(comObj))
            {
                if (ThrowOnRelease)
                    throw new System.Exception("Test release failure");                
                ReleasedObjects.Add(comObj);
            }
        }

        // Override Marshal.IsComObject logic for tests
        protected override bool IsComObjectFunc(object obj)
        {
            return IsComObject;
        }

        public void Set_item(object innerItem)
        { 
            base._item = innerItem;
        }

    }

    public class DummyWithWrongClose
    {
        public void Close() { /* parameterless, will cause TargetParameterCountException */ }
    }

    public interface IFakeOutlookItemNoCloseWithParam
    {
        void Close();
    }
}
