using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gsync.Utilities.Dialogs
{
    public static class InputBox
    {
        public static string ShowDialog(string Prompt, string Title = "", string DefaultResponse = "")
        {
            //if (!InputBoxViewer.DpiCalled) { InputBoxViewer.DpiAware(); }
            var viewer = new InputBoxViewer();
            viewer.AcceptButton = viewer.Ok;
            viewer.CancelButton = viewer.Cancel;
            viewer.Message.Text = Prompt;
            viewer.Text = Title;
            viewer.Input.Text = DefaultResponse;
            viewer.Input.Select();
            DialogResult result = viewer.ShowDialog();
            if (result == DialogResult.OK)
            {
                string value = viewer.Input.Text;
                viewer.Dispose();
                return value;
            }
            else
            {
            viewer.Dispose();
            return null;
            }
        }
    }
}
