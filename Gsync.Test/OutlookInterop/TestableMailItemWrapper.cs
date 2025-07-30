using Gsync.OutlookInterop.Item;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Immutable;

namespace Gsync.Test.OutlookInterop.Item
{
    // Used to verify AttachMailItemEvents gets called
    public class TestableMailItemWrapper : MailItemWrapper
    {
        public bool AttachMailItemEventsCalled { get; private set; }

        public TestableMailItemWrapper(object item)
            : base(item) { }

        public TestableMailItemWrapper(OutlookItemWrapper baseWrapper)
            : base(baseWrapper) { }

        protected override void AttachMailItemEvents()
        {
            AttachMailItemEventsCalled = true;
        }

        // Expose protected constructor for testing
        public TestableMailItemWrapper(object item, ItemEvents_10_Event comEvents, ImmutableHashSet<string> supportedTypes)
            : base(item, comEvents, supportedTypes)
        { }

        public static new ItemEvents_10_Event GetCom10Events(OutlookItemWrapper wrapper) =>
            MailItemWrapper.GetCom10Events(wrapper);
    }
}
