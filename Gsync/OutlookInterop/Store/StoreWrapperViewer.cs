using System;
using System.Windows.Forms;

namespace Gsync.OutlookInterop
{
    public partial class StoreWrapperViewer : Form, IStoreWrapperViewer
    {
        public StoreWrapperViewer()
        {
            InitializeComponent();
            ButtonOk.Click += ButtonOk_Click;
            ButtonCancel.Click += ButtonCancel_Click;
            DisplayName.SelectedValueChanged += DisplayName_SelectedValueChanged;
            ArchiveFS.Click += ArchiveFS_Click;
            ArchiveOutlook.Click += ArchiveOutlook_Click;
            JunkEmail.Click += JunkEmail_Click;
            JunkPotential.Click += JunkPotential_Click;
        }

        public StoreWrapperViewer(StoreWrapperController controller): this()
        {
            Controller = controller;
        }

        public StoreWrapperController Controller { get; set; }

        #region Events

        public void ButtonOk_Click(object sender, EventArgs e)
        {
            Controller?.ButtonOk_Click();
        }

        public void ButtonCancel_Click(object sender, EventArgs e)
        {
            Controller?.ButtonCancel_Click();
        }

        public void DisplayName_SelectedValueChanged(object sender, EventArgs e)
        {
            Controller?.DisplayName_SelectedValueChanged(sender, e);
        }

        public void ArchiveFS_Click(object sender, EventArgs e)
        {
            Controller?.ArchiveFS_Click();
        }

        public void ArchiveOutlook_Click(object sender, EventArgs e)
        {
            Controller?.ArchiveOutlook_Click();
        }

        public void JunkEmail_Click(object sender, EventArgs e)
        {
            Controller?.JunkEmail_Click();
        }

        public void JunkPotential_Click(object sender, EventArgs e)
        {
            Controller?.JunkPotential_Click();
        }

        #endregion Events

        #region Make testable

        public Label JunkEmail { get => _junkEmail; set => _junkEmail = value; }
        public Label JunkPotential { get => _junkPotential; set => _junkPotential = value; }
        public Label ArchiveFS { get => _archiveFS; set => _archiveFS = value; }
        public Label ArchiveOutlook { get => _archiveOutlook; set => _archiveOutlook = value; }
        public Label UserEmail { get => _userEmail; set => _userEmail = value; }
        public Label RootFolder { get => _rootFolder; set => _rootFolder = value; }
        public Label Inbox { get => _inbox; set => _inbox = value; }
        public Button ButtonOk { get => _buttonOk; set => _buttonOk = value; }
        public Button ButtonCancel { get => _buttonCancel; set => _buttonCancel = value; }
        public ComboBox DisplayName { get => _displayName; set => _displayName = value; }

        #endregion Make testable
    }
}
