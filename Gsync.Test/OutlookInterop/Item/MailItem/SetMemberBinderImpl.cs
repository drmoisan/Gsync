using System.Dynamic;

namespace Gsync.Test.OutlookInterop.Item
{
    public class SetMemberBinderImpl : SetMemberBinder
    {
        public SetMemberBinderImpl(string name) : base(name, false) { }
        public override DynamicMetaObject FallbackSetMember(DynamicMetaObject target, DynamicMetaObject value, DynamicMetaObject errorSuggestion) => null;
    }
}
