using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gsync.Utilities.Dialogs
{
    public partial class FolderNotFoundViewer : Form
    {
        public FolderNotFoundViewer()
        {
            InitializeComponent();
        }

        public string FolderAction { get; set; }

        public string FolderName 
        {
            get 
            {
                return FolderNameTxtBox.Text; 
            }
            set
            {
                FolderNameTxtBox.Text = value;
            }
        }

        private void CreateFolder_Click(object sender, EventArgs e)
        {
            FolderAction = "Create";
            this.Hide();
        }

        private void OpenFolder_Click(object sender, EventArgs e)
        {
            FolderAction = "Find";
            this.Hide();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            FolderAction = "Cancel";
            this.Hide();
        }

        private void NoToAll_Click(object sender, EventArgs e)
        {
            FolderAction = "NoToAll";
            this.Hide();
        }
    }
}
