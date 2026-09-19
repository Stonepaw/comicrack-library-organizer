using System;
using System.Drawing;
using System.Windows.Forms;
using LibraryOrganizer.Data;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Dialog
{
    public partial class ConfigureForm : Form
    {
        private readonly ProfileViewModel _profile;

        public ConfigureForm(Profile profile)
        {
            InitializeComponent();

            _profile = new ProfileViewModel(profile);

            profileBindingSource.DataSource = _profile;

            removeEmptyFolderExclusions.SelectedIndex = -1;
        }

        private void failOperationOnEmptyValueDestinationFolderBrowse_Click(
            object sender,
            EventArgs e
        )
        {
            FolderBrowserDialog openFolderDialog = new FolderBrowserDialog();

            if (
                openFolderDialog.ShowDialog(this) == DialogResult.OK
                && openFolderDialog.SelectedPath != null
            )
            {
                _profile.FailOperationOnEmptyValueDestinationFolder = openFolderDialog.SelectedPath;
            }
        }

        private void addEmptyFolderExclusion_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog openFolderDialog = new FolderBrowserDialog();

            if (
                openFolderDialog.ShowDialog(this) == DialogResult.OK
                && openFolderDialog.SelectedPath != null
            )
            {
                removeEmptyFoldersExclusionsBindingSource.Add(openFolderDialog.SelectedPath);
            }
        }

        private void removeEmptyFolderExclusion_Click(object sender, EventArgs e)
        {
            if (removeEmptyFolderExclusions.SelectedIndex >= 0)
            {
                removeEmptyFoldersExclusionsBindingSource.RemoveAt(
                    removeEmptyFolderExclusions.SelectedIndex
                );
            }
        }

        private void removeEmptyFolderExclusions_EnabledChanged(object sender, EventArgs e)
        {
            if (!removeEmptyFolderExclusions.Enabled)
            {
                removeEmptyFolderExclusions.ClearSelected();
            }
        }

        private void failOperationOnEmptyValueFields_EnabledChanged(object sender, EventArgs e)
        {
            if (failOperationOnEmptyValueFields.Enabled)
            {
                failOperationOnEmptyValueFields.DefaultCellStyle.ForeColor =
                    SystemColors.ControlText;
                failOperationOnEmptyValueFields.DefaultCellStyle.SelectionForeColor =
                    SystemColors.ControlText;
            }
            else
            {
                failOperationOnEmptyValueFields.DefaultCellStyle.ForeColor = SystemColors.GrayText;
                failOperationOnEmptyValueFields.DefaultCellStyle.SelectionForeColor =
                    SystemColors.GrayText;
            }
        }
    }
}
