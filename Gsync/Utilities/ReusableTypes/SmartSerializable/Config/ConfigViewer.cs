using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gsync.Utilities.Interfaces;

namespace Gsync.Utilities.ReusableTypes
{
    public partial class ConfigViewer : Form
    {
        #region ctor

        public ConfigViewer()
        {
            InitializeComponent();
            SetupConfigGroupBoxReferences();
        }
        
        private void SetupConfigGroupBoxReferences()
        {
            groupBoxLocal.SpecialFolderComboBox = ComboSpecialFolderLocal;
            groupBoxLocal.RelativePathTextBox = RelativePathLocal;
            groupBoxLocal.FileNameTextBox = FileNameLocal;
            groupBoxLocal.LabelActive = LocalActiveLabel;
            groupBoxLocal.DiskType = ISmartSerializableConfig.ActiveDiskEnum.Local;
            groupBoxLocal.IsActive = true;

            groupBoxNet.SpecialFolderComboBox = ComboSpecialFolderNet;
            groupBoxNet.RelativePathTextBox = RelativePathNet;
            groupBoxNet.FileNameTextBox = FileNameNet;
            groupBoxNet.LabelActive = NetActiveLabel;
            groupBoxNet.DiskType = ISmartSerializableConfig.ActiveDiskEnum.Net;
            groupBoxNet.IsActive = false;

            Boxes = [groupBoxLocal, groupBoxNet];
        }

        public ConfigViewer SetController(ConfigController controller)
        {
            Controller = controller;
            return this;
        }

        #endregion ctor

        #region Properties

        internal List<ConfigGroupBox> Boxes { get; private set; } 
        
        internal ConfigController Controller { get; set; }

        #endregion Properties

        #region Events

        private async void ButtonSave_Click(object sender, EventArgs e) => await Controller?.SaveAsync();

        private void ButtonCancel_Click(object sender, EventArgs e) => Controller?.Cancel();

        private async void ButtonOpen_Click(object sender, EventArgs e) => await Controller?.OpenFileChooserAsync();

        private void GroupBox_Enter(object sender, EventArgs e)
        {
            var gb = (ConfigGroupBox)sender;
            if (!gb.IsActive)
            {
                gb.BackColor = System.Drawing.SystemColors.MenuHighlight;
                gb.ForeColor = System.Drawing.SystemColors.HighlightText;
            }
        }

        private void GroupBox_Click(object sender, EventArgs e)
        {
            var box = (ConfigGroupBox)sender;
            if (!box.IsActive)
            {
                Controller?.ActivateDiskGroup(box.DiskType);
            }

        }

        private void GroupBox_Leave(object sender, EventArgs e)
        {
            var gb = (ConfigGroupBox)sender;
            if (!gb.IsActive)
            {
                gb.BackColor = System.Drawing.SystemColors.Control;
                gb.ForeColor = System.Drawing.SystemColors.ControlText;
            }
        }

        private void SpecialFolder_SelectedValueChanged(object sender, EventArgs e)
        {
            var box = (ConfigGroupBox)((ComboBox)sender).Parent;            
            Controller?.ChangeSpecialFolder(box.SpecialFolderName, box.RelativePath, box.DiskType);
        }

        #endregion Events

        #region Methods

        public void ActivateUiBox(ISmartSerializableConfig.ActiveDiskEnum diskType)
        {
            foreach (var box in Boxes)
            {
                if (box.DiskType == diskType)
                {
                    if (!box.IsActive) { ActivateUiBox(box); }
                }
                else
                {
                    if (box.IsActive) { DeactivateUiBox(box); }
                }
            }
        }

        internal void ActivateUiBox(ConfigGroupBox box) 
        {
            box.IsActive = true;
            box.BackColor = System.Drawing.SystemColors.Highlight;
            box.ForeColor = System.Drawing.SystemColors.HighlightText;
            box.LabelActive.Visible = true;
        }

        internal void DeactivateUiBox(ConfigGroupBox box) 
        {
            box.IsActive = false;
            box.BackColor = System.Drawing.SystemColors.Control;
            box.ForeColor = System.Drawing.SystemColors.ControlText;
            box.LabelActive.Visible = false;
        }

        #endregion Methods

    }
}
