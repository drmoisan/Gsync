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
    public partial class InputBoxViewer : Form
    {
        public InputBoxViewer()
        {
            InitializeComponent();
        }

        [STAThread]
        public static void DpiAware()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            DpiCalled = true;
        }

        public static bool DpiCalled { get; set; } = false;

        private void Ok_Click(object sender, EventArgs e)
        {
            if (Input.Text == "") 
            {
                MessageBox.Show("Please enter a value or cancel.");
            }
            else 
            {  
                DialogResult = DialogResult.OK;
                this.Hide();
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Hide();
        }
    }
}
