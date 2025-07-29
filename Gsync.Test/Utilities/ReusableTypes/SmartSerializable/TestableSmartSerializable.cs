using Gsync.Test.Utilities.ReusableTypes.SmartSerializable;
using Gsync.Utilities.HelperClasses;
using Gsync.Utilities.ReusableTypes;

namespace Gsync.Test.Utilities.ReusableTypes
{
    public class TestableSmartSerializable : SmartSerializable<TestConfig>
    {
        public TestableSmartSerializable(TestConfig parent) : base(parent) { }

        public TestConfig CallDeserializeJson(FilePathHelper disk)
            => base.DeserializeJson(disk);
    }

}
