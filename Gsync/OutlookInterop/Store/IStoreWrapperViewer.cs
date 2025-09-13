using System;
using System.Windows.Forms;
using Gsync.OutlookInterop.Interfaces.IWinForm;

namespace Gsync.OutlookInterop
{
    public interface IStoreWrapperViewer: IForm
    {
        Label ArchiveFS { get; set; }
        Label ArchiveOutlook { get; set; }
        Button ButtonCancel { get; set; }
        Button ButtonOk { get; set; }
        ComboBox DisplayName { get; set; }
        Label Inbox { get; set; }
        Label JunkEmail { get; set; }
        Label JunkPotential { get; set; }
        Label RootFolder { get; set; }
        Label UserEmail { get; set; }

        void ButtonCancel_Click(object sender, EventArgs e);
        void ButtonOk_Click(object sender, EventArgs e);
        void DisplayName_SelectedValueChanged(object sender, EventArgs e);
        public void ArchiveFS_Click(object sender, EventArgs e);
        public void ArchiveOutlook_Click(object sender, EventArgs e);
        public void JunkEmail_Click(object sender, EventArgs e);
        public void JunkPotential_Click(object sender, EventArgs e);
    }
}