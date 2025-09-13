using System.Windows.Forms;

namespace Gsync.Utilities.Interfaces
{
    public interface IUserDialog
    {
        DialogResult ShowDialog(string message, string caption, MessageBoxButtons buttons, MessageBoxIcon icon);
    }
}
