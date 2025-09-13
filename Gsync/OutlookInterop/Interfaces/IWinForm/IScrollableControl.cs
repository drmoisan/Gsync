using System.Drawing;
using System.Windows.Forms;

namespace Gsync.OutlookInterop.Interfaces.IWinForm
{ 
    public interface IScrollableControl: IControl
    {
        bool AutoScroll { get; set; }
        Size AutoScrollMargin { get; set; }
        Size AutoScrollMinSize { get; set; }
        Point AutoScrollPosition { get; set; }        
        ScrollableControl.DockPaddingEdges DockPadding { get; }
        HScrollProperties HorizontalScroll { get; }
        VScrollProperties VerticalScroll { get; }

        event ScrollEventHandler Scroll;

        void ScrollControlIntoView(Control activeControl);
        void SetAutoScrollMargin(int x, int y);
    }
}