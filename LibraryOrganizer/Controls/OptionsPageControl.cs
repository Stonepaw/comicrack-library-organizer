using System;
using System.Windows.Forms;
using LibraryOrganizer.Data;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class OptionsPageControl : UserControl
    {
        private ProfileViewModel _profileViewModel;

        public ProfileViewModel ProfileViewModel
        {
            get => _profileViewModel;
            set
            {
                if (value == null)
                {
                    return;
                }

                _profileViewModel = value;
                profileViewModelBindingSource.DataSource = _profileViewModel;
            }
        }

        public OptionsPageControl()
        {
            InitializeComponent();
            removeEmptyFolderExclusions.SelectedIndex = -1;
        }

        private void addEmptyFolderExclusion_Click(object sender, EventArgs e)
        {
            string folder = FolderBrowserUtil.PromptForFolderPath(this);

            if (folder == null)
            {
                return;
            }

            removeEmptyFoldersExclusionsBindingSource.Add(folder);
        }

        private void addIllegalCharacterReplacement_Click(object sender, EventArgs e)
        {
            var addIllegalCharacterDialog = new AddIllegalCharacterDialog(
                _profileViewModel.IllegalCharacterReplacements
            );

            if (addIllegalCharacterDialog.ShowDialog(this) == DialogResult.OK)
            {
                var index = illegalCharacterReplacementsBindingSource.Add(
                    new IllegalCharacterReplacement(addIllegalCharacterDialog.GetCharacter(), "")
                );
                illegalCharacterReplacementsBindingSource.Position = index;
            }
        }

        private void removeEmptyFolderExclusions_EnabledChanged(object sender, EventArgs e)
        {
            if (!removeEmptyFolderExclusions.Enabled)
            {
                removeEmptyFolderExclusions.ClearSelected();
            }
        }

        private void removeEmptyFolderExclusion_Click(object sender, EventArgs e)
        {
            removeEmptyFoldersExclusionsBindingSource.RemoveCurrent();
        }

        private void illegalCharacterReplacementsReplacement_KeyPress(
            object sender,
            KeyPressEventArgs e
        )
        {
            if (IllegalCharacterReplacement.IsRequiredIllegalCharacter(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void removeIllegalCharacterReplacement_Click(object sender, EventArgs e)
        {
            if (
                !(
                    (IllegalCharacterReplacement)illegalCharacterReplacementsBindingSource.Current
                ).IsRequired()
            )
            {
                illegalCharacterReplacementsBindingSource.RemoveCurrent();
            }
        }
    }
}
