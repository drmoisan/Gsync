using Newtonsoft.Json;
using System;
using System.ComponentModel;
using Gsync.Utilities.ReusableTypes;
using Gsync.Utilities.HelperClasses;

namespace Gsync.Utilities.Interfaces
{
    public interface ISmartSerializableConfig: ICloneable, INotifyPropertyChanged
    {
        bool ClassifierActivated { get; set; }
        FilePathHelper Disk { get; set; }
        JsonSerializerSettings JsonSettings { get; set; }
        DateTime LocalDate { get; }
        FilePathHelper LocalDisk { get; set; }
        JsonSerializerSettings LocalJsonSettings { get; set; }
        FilePathHelper NetDisk { get; set; }
        JsonSerializerSettings NetJsonSettings { get; set; }
        DateTime NetworkDate { get; }
        ActiveDiskEnum ActiveDisk { get; }
        ISmartSerializableConfig DeepCopy();
        void CopyFrom(ISmartSerializableConfig config, bool deep);
        void CopyChanged(ISmartSerializableConfig other, bool deep, bool notify);

        void ActivateLocalDisk();
        void ActivateMostRecent();
        void ActivateNetDisk();
        void ResetLazy();
        void ResetLazy(Lazy<JsonSerializerSettings> localJsonSettings, Lazy<JsonSerializerSettings> netJsonSettings, Lazy<JsonSerializerSettings> jsonSettings);

        enum ActiveDiskEnum
        {
            Neither,
            Local,
            Net
        }
    }
}