using System;
using System.Collections.Generic;
using System.Dynamic;

namespace Gsync.Test.OutlookInterop.Item
{
    public class DynamicIItem : DynamicObject
    {
        public Dictionary<string, object> Properties = new();
        public Dictionary<string, Delegate> Methods = new();

        public override bool TryGetMember(GetMemberBinder binder, out object result)
            => Properties.TryGetValue(binder.Name, out result);

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            Properties[binder.Name] = value;
            return true;
        }

        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            if (Methods.TryGetValue(binder.Name, out var del))
            {
                result = del.DynamicInvoke(args);
                return true;
            }
            result = null;
            return false;
        }
    }
}
