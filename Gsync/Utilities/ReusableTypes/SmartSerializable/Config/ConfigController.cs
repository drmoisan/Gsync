using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gsync.Utilities;
using Gsync.Utilities.HelperClasses.NewtonSoft;
using Gsync.Utilities.Interfaces;


namespace Gsync.Utilities.ReusableTypes
{
    public class ConfigController
    {

        #region ctor

        public ConfigController(AppGlobals globals, ISmartSerializableConfig config) 
        { 
            Config = config;
            ConfigCopy = config.DeepCopy();
            Globals = globals;           
        }

        public ConfigController Init() 
        {
            FilePathConverter = new FilePathHelperConverter(Globals.FS);
            SpecialFolderList = [.. Globals.FS.SpecialFolders.Keys];
            SpecialFolderList.Insert(0, "None");

            Viewer = new ConfigViewer();            
            
            Viewer.ComboSpecialFolderLocal.DataSource = Globals.FS.SpecialFolders.Keys;
            var (specialFolderLocal, relativePathLocal) = FilePathConverter.GetSerializablePath(ConfigCopy.LocalDisk.FolderPath);
            var specialFolderSelectionLocal = specialFolderLocal == "Not Found" ? "None" : specialFolderLocal;
            Viewer.ComboSpecialFolderLocal.SelectedItem = specialFolderSelectionLocal;
            Viewer.RelativePathLocal.Text = relativePathLocal;
            Viewer.FileNameLocal.Text = ConfigCopy.LocalDisk.FileName;

            Viewer.ComboSpecialFolderNet.DataSource = Globals.FS.SpecialFolders.Keys;
            var (specialFolderNet, relativePathNet) = FilePathConverter.GetSerializablePath(ConfigCopy.NetDisk.FolderPath);
            var specialFolderSelectionNet = specialFolderNet == "Not Found" ? "None" : specialFolderNet;
            Viewer.ComboSpecialFolderNet.SelectedItem = specialFolderSelectionNet;
            Viewer.RelativePathNet.Text = relativePathNet;
            Viewer.FileNameNet.Text = ConfigCopy.NetDisk.FileName;

            Viewer.ActivateUiBox(ConfigCopy.ActiveDisk);

            Viewer.SetController(this);
            return this;
        }

        public static ConfigController Show(AppGlobals globals, ISmartSerializableConfig config)
        {            
            var controller = new ConfigController(globals, config).Init();            
            controller.Viewer.Show();
            return controller;
        }

        #endregion ctor

        #region Properties
        internal ISmartSerializableConfig ConfigCopy { get; set; }
        internal ISmartSerializableConfig Config { get; set; }
        internal ConfigViewer Viewer { get; set; }
        internal AppGlobals Globals { get; set; }
        internal FilePathHelperConverter FilePathConverter { get; set; }
        internal List<string> SpecialFolderList { get; set; }
        #endregion Properties

        #region Events

        internal void Cancel()
        {
            Viewer.Close();
        }

        internal void ChangeSpecialFolder(string specialFolderName, string relativePath, ISmartSerializableConfig.ActiveDiskEnum diskType)
        {
            var folderPath = FilePathConverter.ExtractFolderPath(specialFolderName, relativePath);
            if (ConfigCopy.ActiveDisk == diskType) { ConfigCopy.Disk.FolderPath = folderPath; }
            
            switch (diskType)
            {
                case ISmartSerializableConfig.ActiveDiskEnum.Local:
                    ConfigCopy.LocalDisk.FolderPath = folderPath; // FilePathConverter.ExtractFolderPath(specialFolderName, relativePath);
                    break;
                case ISmartSerializableConfig.ActiveDiskEnum.Net:
                    ConfigCopy.NetDisk.FolderPath = folderPath; // FilePathConverter.ExtractFolderPath(specialFolderName, relativePath);
                    break;
                default:
                    break;
            }            
        }

        internal void ActivateDiskGroup(ISmartSerializableConfig.ActiveDiskEnum diskType)
        {
            switch (diskType)
            {
                case ISmartSerializableConfig.ActiveDiskEnum.Local:
                    ConfigCopy.ActivateLocalDisk();
                    Viewer.ActivateUiBox(diskType);
                    break;
                case ISmartSerializableConfig.ActiveDiskEnum.Net:
                    ConfigCopy.ActivateNetDisk();
                    Viewer.ActivateUiBox(diskType);
                    break;
                default:
                    break;
            }
        }

        internal async Task OpenFileChooserAsync()
        {
            await Task.CompletedTask;
            throw new NotImplementedException();
        }

        internal async Task SaveAsync()
        {
            if (SynchronizationContext.Current is null)
                SynchronizationContext.SetSynchronizationContext(
                    new WindowsFormsSynchronizationContext());
            Viewer.Enabled = false;
            await Task.Run(() => Config.CopyChanged(ConfigCopy, true, true));

            Viewer.Close();
        }

        #endregion Events

        
    }
}
