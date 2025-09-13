using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gsync.Utilities.Interfaces;

namespace Gsync.Utilities.ReusableTypes
{
    internal class ConfigGroupBox:GroupBox
    {
        internal ISmartSerializableConfig.ActiveDiskEnum DiskType { get; set; }

        internal TextBox FileNameTextBox { get; set; }
        internal string FileName { get => FileNameTextBox.Text; set => FileNameTextBox.Text = value; }

        internal bool IsActive { get; set; }

        internal Label LabelActive { get; set; }

        internal TextBox RelativePathTextBox { get; set; }
        internal string RelativePath { get => RelativePathTextBox.Text; set => RelativePathTextBox.Text = value; }
        
        internal ComboBox SpecialFolderComboBox {  get; set; }
        internal string SpecialFolderName 
        {
            get => SpecialFolderComboBox.SelectedItem as string;
            set => SpecialFolderComboBox.SelectedItem = SpecialFolderComboBox.Items.Contains(value) ? value : null;            
        }


    }
}
