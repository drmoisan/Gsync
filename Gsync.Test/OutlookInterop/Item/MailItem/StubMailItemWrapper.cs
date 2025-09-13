using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;


namespace Gsync.OutlookInterop.Item
{
    public class StubMailItemWrapper : MailItemWrapper
    {
        public StubMailItemWrapper(object item, ItemEvents_10_Event comEvents, ImmutableHashSet<string> supportedTypes)
            : base(item, comEvents, supportedTypes) { }
    }
    //public class StubOutlookItemWrapper : OutlookItemWrapper
    //{
    //    public StubOutlookItemWrapper(object item, ItemEvents_10_Event comEvents, ImmutableHashSet<string> supportedTypes)
    //        : base(item, comEvents, supportedTypes) { }
    //}

}
