using Newtonsoft.Json;
using System;
using System.IO;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using Gsync.Utilities.Interfaces;
using Gsync.Utilities.Extensions;
using Gsync.Utilities.HelperClasses;

namespace Gsync.Utilities.ReusableTypes
{
    public class NewSmartSerializableConfig : ISmartSerializableConfig
    {
        public NewSmartSerializableConfig()
        {
            ResetLazy();
            _activeDisk = ISmartSerializableConfig.ActiveDiskEnum.Neither;
        }

        #region SerializationConfig

        protected FilePathHelper _disk = new FilePathHelper();
        public FilePathHelper Disk { get => _disk; set => _disk = value; }

        public FilePathHelper LocalDisk { get => _localDisk; set => _localDisk = value; }
        private FilePathHelper _localDisk = new FilePathHelper();

        public FilePathHelper NetDisk { get => _netDisk; set => _netDisk = value; }
        private FilePathHelper _netDisk = new();

        [JsonIgnore]
        public DateTime NetworkDate => File.Exists(NetDisk.FilePath) ?
            File.GetLastWriteTimeUtc(NetDisk.FilePath) : default;

        [JsonIgnore]
        public DateTime LocalDate => File.Exists(LocalDisk.FilePath) ?
            File.GetLastWriteTimeUtc(LocalDisk.FilePath) : default;

        private bool _classifierActivated;
        public bool ClassifierActivated { get => _classifierActivated; set { _classifierActivated = value; Notify(); } }

        public void ResetLazy()
        {
            _localJsonSettings = new Lazy<JsonSerializerSettings>(GetDefaultSettings);
            _netJsonSettings = new Lazy<JsonSerializerSettings>(GetDefaultSettings);
            _jsonSettings = new Lazy<JsonSerializerSettings>(GetDefaultSettings);
        }

        public void ResetLazy(
            Lazy<JsonSerializerSettings> localJsonSettings,
            Lazy<JsonSerializerSettings> netJsonSettings,
            Lazy<JsonSerializerSettings> jsonSettings)
        {
            _localJsonSettings = localJsonSettings;
            _netJsonSettings = netJsonSettings;
            _jsonSettings = jsonSettings;
        }

        [JsonIgnore]
        public JsonSerializerSettings JsonSettings { get => _jsonSettings.Value; set { _jsonSettings = value.ToLazy(); Notify(); } }
        protected Lazy<JsonSerializerSettings> _jsonSettings;

        [JsonIgnore]
        public JsonSerializerSettings NetJsonSettings { get => _netJsonSettings.Value; set { _netJsonSettings = value.ToLazy(); Notify(); } }
        protected Lazy<JsonSerializerSettings> _netJsonSettings;

        [JsonIgnore]
        public JsonSerializerSettings LocalJsonSettings { get => _localJsonSettings.Value; set { _localJsonSettings = value.ToLazy(); Notify(); } }
        protected Lazy<JsonSerializerSettings> _localJsonSettings;

        public ISmartSerializableConfig.ActiveDiskEnum ActiveDisk { get => _activeDisk; protected set { _activeDisk = value; Notify(); } } 
        protected ISmartSerializableConfig.ActiveDiskEnum _activeDisk;

        public static JsonSerializerSettings GetDefaultSettings()
        {
            return new JsonSerializerSettings()
            {
                TypeNameHandling = TypeNameHandling.Auto,
                Formatting = Formatting.Indented
            };
        }

        public void ActivateMostRecent()
        {
            if (NetworkDate != default && (LocalDate == default || NetworkDate > LocalDate))
            {
                ActivateNetDisk();
            }
            else
            {
                ActivateLocalDisk();
            }
        }

        public void ActivateLocalDisk()
        {
            _disk.CopyFrom(_localDisk);
            _jsonSettings = _localJsonSettings;
            ActiveDisk = ISmartSerializableConfig.ActiveDiskEnum.Local;
        }

        public void ActivateNetDisk()
        {
            _disk.CopyFrom(_netDisk);
            _jsonSettings = _netJsonSettings;
            ActiveDisk = ISmartSerializableConfig.ActiveDiskEnum.Net;
        }

        #endregion SerializationConfig

        #region IClonable

        public object Clone()
        {
            return this.MemberwiseClone();
        }

        public ISmartSerializableConfig DeepCopy() 
        {
            //var clone = (NewSmartSerializableConfig)Clone();
            var clone = new NewSmartSerializableConfig
            {
                _activeDisk = _activeDisk,
                _classifierActivated = _classifierActivated,
                _disk = _disk.DeepCopy(),
                _localDisk = _localDisk.DeepCopy(),
                _netDisk = _netDisk.DeepCopy(),
                _jsonSettings = JsonSettings.DeepCopy().ToLazy(),
                _netJsonSettings = NetJsonSettings.DeepCopy().ToLazy(),
                _localJsonSettings = LocalJsonSettings.DeepCopy().ToLazy()
            };
            return clone;
        }

        public void CopyFrom(ISmartSerializableConfig other, bool deep)
        {
            if (deep) { other = other.DeepCopy(); }

            // Using private fields to avoid triggering events recursively
            _activeDisk = other.ActiveDisk;
            _classifierActivated = other.ClassifierActivated;
            Disk.CopyFrom(other.Disk);
            LocalDisk.CopyFrom(other.LocalDisk);
            NetDisk.CopyFrom(other.NetDisk);
            _jsonSettings = other.JsonSettings.ToLazy();
            _netJsonSettings = other.NetJsonSettings.ToLazy();
            _localJsonSettings = other.LocalJsonSettings.ToLazy();
            Notify("CopyFrom");
        }

        public void CopyChanged(ISmartSerializableConfig other, bool deep, bool notify) 
        {
            var changed = CopyChanged(other, deep);
            if (notify) 
            { 
                Notify(string.Join(",", changed));
            }
        }

        public IList<string> CopyChanged(ISmartSerializableConfig other, bool deep)
        {
            List<string> changed = [];
            
            if (_activeDisk != other.ActiveDisk)
            {
                _activeDisk = other.ActiveDisk;
                changed.Add(nameof(ActiveDisk));
            }

            if (_classifierActivated != other.ClassifierActivated)
            {
                _classifierActivated = other.ClassifierActivated;
                changed.Add(nameof(ClassifierActivated));
            }

            changed.AddRange(Disk.CopyChanged(other.Disk).Select(x => $"{nameof(Disk)}.{x}"));
            changed.AddRange(LocalDisk.CopyChanged(other.LocalDisk).Select(x => $"{nameof(LocalDisk)}.{x}"));
            changed.AddRange(NetDisk.CopyChanged(other.NetDisk).Select(x => $"{nameof(NetDisk)}.{x}"));

            if (JsonSettings != other.JsonSettings)
            {
                _jsonSettings = other.JsonSettings.ToLazy();
                changed.Add(nameof(JsonSettings));
            }

            if (NetJsonSettings != other.NetJsonSettings)
            {
                _netJsonSettings = other.NetJsonSettings.ToLazy();
                changed.Add(nameof(NetJsonSettings));
            }

            if (LocalJsonSettings != other.LocalJsonSettings)
            {
                _localJsonSettings = other.LocalJsonSettings.ToLazy();
                changed.Add(nameof(LocalJsonSettings));
            }

            return changed;


        }

        #endregion IClonable

        #region INotifyPropertyChanged

        public void Notify([System.Runtime.CompilerServices.CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion INotifyPropertyChanged


    }
}
