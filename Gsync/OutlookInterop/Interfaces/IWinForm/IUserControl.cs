
using System;
using System.Windows.Forms;

namespace Gsync.OutlookInterop.Interfaces.IWinForm
{
    public interface IUserControl: IContainerControl
    {
        bool AutoSize { get; set; }
        AutoSizeMode AutoSizeMode { get; set; }
        AutoValidate AutoValidate { get; set; }
        BorderStyle BorderStyle { get; set; }
        string Text { get; set; }

        event EventHandler AutoSizeChanged;
        event EventHandler AutoValidateChanged;
        event EventHandler Load;
        event EventHandler TextChanged;

        bool ValidateChildren();
        bool ValidateChildren(ValidationConstraints validationConstraints);
    }
}