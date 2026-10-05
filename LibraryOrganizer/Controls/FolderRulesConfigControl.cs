using System;
using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class FolderRulesConfigControl : UserControl
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

        public FolderRulesConfigControl()
        {
            InitializeComponent();
        }

        private void addExcludedFolder_Click(object sender, EventArgs e)
        {
            string folder = FolderBrowserUtil.PromptForFolderPath(this);

            if (folder != null)
            {
                excludeFoldersBindingSource.Add(folder);
            }
        }

        private void removeExcludedFolder_Click(object sender, EventArgs e)
        {
            excludeFoldersBindingSource.RemoveCurrent();
        }
    }
}
