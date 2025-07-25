using System;
using System.Linq;
using System.Windows.Forms;
using Gsync.Utilities.HelperClasses;
using Gsync.Utilities.Interfaces;
using Gsync.Utilities.HelperClasses.NewtonSoft;
using Gsync.Utilities.Extensions;
using Gsync.Utilities.Dialogs;


namespace Gsync.OutlookInterop
{
    public class StoreWrapperController
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #region ctor

        public StoreWrapperController(IApplicationGlobals globals)
        {
            Globals = globals;            
        }

        #endregion ctor

        internal IApplicationGlobals Globals { get; set; }
        
        public IStoreWrapperViewer Viewer { get; internal set; }

        public StoresWrapper Model { get; internal set; }

        public StoreWrapper Current { get; internal set; }

        internal FolderMinimalWrapper ArchiveOutlook { get; set; }
        internal FilePathHelper ArchiveFS { get; set; }
        internal FolderMinimalWrapper JunkEmail { get; set; }
        internal FolderMinimalWrapper JunkPotential { get; set; }       
        internal Func<string, (string, string)> FsConverter { get; set; }
        

        #region Events

        public void Launch() 
        {
            FsConverter = new FilePathHelperConverter(Globals.FS).GetSerializablePath;
            Model = Globals.Stores;
            Viewer = new StoreWrapperViewer(this);
            Viewer.DisplayName.DataSource = Model.Stores.Select(store => store.DisplayName).ToList();
            
            Viewer.ShowDialog();
        }

        public void ButtonOk_Click()
        {
            if (AnyChanges()) { SaveChanges(); }
            Viewer.Close();
        }

        public void ButtonCancel_Click()
        {
            Viewer.Close();
        }

        public void DisplayName_SelectedValueChanged(object sender, EventArgs e)
        {
            if (AnyChanges())
            {
                var response = MyBox.ShowDialog("Save changes?", "Save Changes", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (response == DialogResult.Yes) { SaveChanges(); }
            }
            var displayName = Viewer.DisplayName.SelectedValue.ToString();
            Current = Model.Stores.Find(store => store.DisplayName == displayName);
            PopulateWithCurrent();
        }

        public void ArchiveFS_Click()
        {
            if (Viewer.InvokeRequired)
            {
                Viewer.Invoke(() => ArchiveFS_Click());
                return;
            }
            var folderPath = SelectFsFolder();
            if (folderPath.IsNullOrEmpty()) { return; }
            else 
            {                 
                ArchiveFS.FolderPath = folderPath;
                Viewer.ArchiveFS.Text = GetRelativeFsPath();
            }
        }

        public void ArchiveOutlook_Click()
        {
            if (Viewer.InvokeRequired)
            {
                Viewer.Invoke(() => ArchiveOutlook_Click());
                return;
            }
            ArchiveOutlook = SelectFolder();
            Viewer.ArchiveOutlook.Text = ArchiveOutlook?.RelativePath;
        }

        public void JunkEmail_Click()
        {
            if (Viewer.InvokeRequired)
            {
                Viewer.Invoke(() => JunkEmail_Click());
                return;
            }
            JunkEmail = SelectFolder();
            Viewer.JunkEmail.Text = JunkEmail?.RelativePath;
        }

        public void JunkPotential_Click()
        {
            if (Viewer.InvokeRequired)
            {
                Viewer.Invoke(() => JunkPotential_Click());
                return;
            }
            JunkPotential = SelectFolder();
            Viewer.JunkPotential.Text = JunkPotential?.RelativePath;
        }

        #endregion Events

        #region Methods
        
        internal bool AnyChanges()
        {
            return !PairwiseEquals(ArchiveOutlook, Current?.ArchiveRoot) ||
                !PairwiseEquals(JunkEmail, Current?.JunkCertain) ||
                !PairwiseEquals(JunkPotential, Current?.JunkPotential) ||
                !PairwiseEquals(ArchiveFS, Current?.ArchiveFsRoot);
        }

        internal bool PairwiseEquals<T>(T a, T b)
        {
            if (a is null && b is null) { return true; }
            if (a is null || b is null) { return false; }
            return a.Equals(b);
        }

        internal void PopulateWithCurrent()
        {
            if (Viewer.InvokeRequired) { Viewer.Invoke(() => PopulateWithCurrent()); return; }

            // Populate Form
            Viewer.Inbox.Text = Current?.Inbox?.FolderPath ?? "Error Loading";
            Viewer.RootFolder.Text = Current?.RootFolder?.FolderPath ?? "Error Loading";
            Viewer.UserEmail.Text = Current?.UserEmailAddress ?? "Error Loading";
            Viewer.ArchiveOutlook.Text = Current?.ArchiveRoot?.RelativePath ?? "Please select an archive";
            Viewer.ArchiveFS.Text = GetRelativeFsPath();
            //if (Current.ArchiveFsRoot is not null && !Current.ArchiveFsRoot.FolderPath.IsNullOrEmpty())
            //{
            //    var (specialFolder, relativePath) = FsConverter(Current.ArchiveFsRoot.FolderPath);
            //    if (specialFolder.IsNullOrEmpty() & relativePath.IsNullOrEmpty())
            //    {
            //        Viewer.ArchiveFS.Text = "Please select an archive";
            //    }
            //    else
            //    {
            //        Viewer.ArchiveFS.Text = $"{string.Join(" -> ", [specialFolder,relativePath]).Trim()}";
            //    }
            //}
            Viewer.JunkEmail.Text = JunkEmail?.RelativePath ?? "Please select a folder";
            Viewer.JunkPotential.Text = JunkPotential?.RelativePath ?? "Please select a folder";

            // Populate Controller
            ArchiveOutlook = Current.ArchiveRoot;
            ArchiveFS = Current.ArchiveFsRoot;
            JunkEmail = Current.JunkCertain;
            JunkPotential = Current.JunkPotential;
        }

        internal void SaveChanges()
        {            
            Current.ArchiveRoot = ArchiveOutlook;
            Current.JunkCertain = JunkEmail;
            Current.JunkPotential = JunkPotential;
            Current.ArchiveFsRoot = ArchiveFS;
            Model.Serialize();
        }

        internal virtual FolderMinimalWrapper SelectFolder()
        {
            try
            {
                var folder = Globals.OutlookApplication.Session.PickFolder();
                if (folder is null) { return null; }
                return new FolderMinimalWrapper(folder, Current.RootFolder);
            }
            catch (Exception e)
            {
                logger.Error($"Error selecting folder. {e.Message}", e);
                return null;
            }                   
        }

        internal string SelectFsFolder()
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select a folder";
                folderBrowserDialog.ShowNewFolderButton = true;
                folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer;

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of the selected folder
                    return folderBrowserDialog.SelectedPath;
                }
            }
            return null;
        }

        internal string GetRelativeFsPath()
        {
            if (Current.ArchiveFsRoot is not null && !Current.ArchiveFsRoot.FolderPath.IsNullOrEmpty())
            {
                var (specialFolder, relativePath) = FsConverter(Current.ArchiveFsRoot.FolderPath);
                if (specialFolder.IsNullOrEmpty() & relativePath.IsNullOrEmpty())
                {
                    return "Please select an archive";
                }
                else
                {
                    return $"{string.Join(" -> ", [specialFolder, relativePath]).Trim()}";
                }
            }
            return "Please select an archive";
        }

        #endregion Methods

    }
}
