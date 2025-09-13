using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace Gsync.OutlookInterop.Interfaces.IWinForm
{
    public interface IForm: IContainerControl, IScrollableControl
    {
        IButtonControl AcceptButton { get; set; }
        Form ActiveMdiChild { get; }
        bool AllowTransparency { get; set; }
        bool AutoScale { get; set; }
        Size AutoScaleBaseSize { get; set; }                
        AutoSizeMode AutoSizeMode { get; set; }
        AutoValidate AutoValidate { get; set; }        
        IButtonControl CancelButton { get; set; }        
        bool ControlBox { get; set; }
        Rectangle DesktopBounds { get; set; }
        Point DesktopLocation { get; set; }
        DialogResult DialogResult { get; set; }
        FormBorderStyle FormBorderStyle { get; set; }
        bool HelpButton { get; set; }
        Icon Icon { get; set; }
        bool IsMdiChild { get; }
        bool IsMdiContainer { get; set; }
        bool IsRestrictedWindow { get; }
        bool KeyPreview { get; set; }        
        MenuStrip MainMenuStrip { get; set; }        
        bool MaximizeBox { get; set; }        
        Form[] MdiChildren { get; }
        Form MdiParent { get; set; }
        MainMenu Menu { get; set; }
        MainMenu MergedMenu { get; }
        bool MinimizeBox { get; set; }        
        bool Modal { get; }
        double Opacity { get; set; }
        Form[] OwnedForms { get; }
        Form Owner { get; set; }
        Rectangle RestoreBounds { get; }
        bool RightToLeftLayout { get; set; }
        bool ShowIcon { get; set; }
        bool ShowInTaskbar { get; set; }        
        SizeGripStyle SizeGripStyle { get; set; }
        FormStartPosition StartPosition { get; set; }                        
        bool TopLevel { get; set; }
        bool TopMost { get; set; }
        Color TransparencyKey { get; set; }
        FormWindowState WindowState { get; set; }

        event EventHandler Activated;        
        event EventHandler AutoValidateChanged;
        event EventHandler Closed;
        event CancelEventHandler Closing;
        event EventHandler Deactivate;
        event DpiChangedEventHandler DpiChanged;
        event FormClosedEventHandler FormClosed;
        event FormClosingEventHandler FormClosing;
        event CancelEventHandler HelpButtonClicked;
        event InputLanguageChangedEventHandler InputLanguageChanged;
        event InputLanguageChangingEventHandler InputLanguageChanging;
        event EventHandler Load;        
        event EventHandler MaximizedBoundsChanged;
        event EventHandler MaximumSizeChanged;
        event EventHandler MdiChildActivate;
        event EventHandler MenuComplete;
        event EventHandler MenuStart;
        event EventHandler MinimumSizeChanged;
        event EventHandler ResizeBegin;
        event EventHandler ResizeEnd;
        event EventHandler RightToLeftLayoutChanged;
        event EventHandler Shown;                

        void Activate();
        void AddOwnedForm(Form ownedForm);
        void Close();
        void LayoutMdi(MdiLayout value);
        void RemoveOwnedForm(Form ownedForm);
        void SetDesktopBounds(int x, int y, int width, int height);
        void SetDesktopLocation(int x, int y);
        void Show(IWin32Window owner);
        DialogResult ShowDialog();
        DialogResult ShowDialog(IWin32Window owner);
        string ToString();
        bool ValidateChildren();
        bool ValidateChildren(ValidationConstraints validationConstraints);
    }
}