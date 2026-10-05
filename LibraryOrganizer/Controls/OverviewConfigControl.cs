using System.ComponentModel;
using System.Windows.Forms;
using LibraryOrganizer.Data;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class OverviewConfigControl : UserControl
    {
        private ProfileViewModel _profileViewModel;

        public ProfileViewModel ProfileViewModel
        {
            get => _profileViewModel;
            set
            {
                if (value != null)
                {
                    _profileViewModel = value;
                    profileViewModelBindingSource.DataSource = value;
                }
            }
        }

        public OverviewConfigControl()
        {
            _profileViewModel = new ProfileViewModel(new Profile());
            InitializeComponent();
            profileViewModelBindingSource.DataSource = _profileViewModel;
        }

        private void baseFolderBrowse_Click(object sender, System.EventArgs e)
        {
            string folder = FolderBrowserUtil.PromptForFolderPath(this);

            if (folder == null)
            {
                return;
            }

            if (_profileViewModel != null)
            {
                _profileViewModel.BaseFolder = folder;
            }
        }
    }
}
