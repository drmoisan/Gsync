using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gsync.Utilities.Dialogs
{
    public static class NotImplementedDialog
    {
        private delegate DialogResult ResponseDelegate();

        public static bool StopAtNotImplemented(string functionName) 
        {
            string title = "Not Implemented Dialog";
            string message = "Function " + functionName + " is not implemented. Throw exception or keep running?";
            Dictionary<string, Delegate> map = new()
            {
                { "Throw Exception", new ResponseDelegate(ThrowException) },
                { "Keep Running", new ResponseDelegate(KeepRunning) }
            };
            MyBoxViewer _box = new(title, message, map);
            DialogResult result = _box.ShowDialog();
            if (result == DialogResult.Yes) {return true;}
            else { return false; }
        }

        private static DialogResult ThrowException() { return DialogResult.Yes; }
        private static DialogResult KeepRunning() { return DialogResult.No; }
    }
}
