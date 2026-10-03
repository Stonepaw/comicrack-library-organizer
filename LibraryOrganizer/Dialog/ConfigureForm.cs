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

        private readonly ConfigFormViewModel _configFormViewModel = new ConfigFormViewModel();

        public ConfigureForm(Profile profile)
        {
            InitializeComponent();

            SuspendLayout();

            _profile = new ProfileViewModel(profile);

            profileBindingSource.DataSource = _profile;
            configFormViewModelBindingSource.DataSource = _configFormViewModel;

            removeEmptyFolderExclusions.SelectedIndex = -1;
            overviewButton.Tag = ConfigFormPage.Overview;
            filesButton.Tag = ConfigFormPage.Files;
            foldersButton.Tag = ConfigFormPage.Folders;
            rulesButton.Tag = ConfigFormPage.Rules;
            optionsButton.Tag = ConfigFormPage.Options;
            SetCurrentPage(ConfigFormPage.Overview);
            ShowPage(ConfigFormPage.Overview);

            ResumeLayout();
        }

        private void failOperationOnEmptyValueDestinationFolderBrowse_Click(
            object sender,
            EventArgs e
        )
        {
            var openFolderDialog = new FolderBrowserDialog();

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
            var openFolderDialog = new FolderBrowserDialog();

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

        private void addIllegalCharacterReplacement_Click(object sender, EventArgs e)
        {
            var addIllegalCharacterDialog = new AddIllegalCharacterDialog(
                _profile.IllegalCharacterReplacements
            );

            if (addIllegalCharacterDialog.ShowDialog(this) == DialogResult.OK)
            {
                var index = illegalCharacterReplacementsBindingSource.Add(
                    new IllegalCharacterReplacement(addIllegalCharacterDialog.GetCharacter(), "")
                );
                illegalCharacterReplacementsBindingSource.Position = index;
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

        /// <summary>
        /// Disallows entering required illegal characters since they would just get replaced anyway.
        /// </summary>
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

        private void PageButton_Click(object sender, EventArgs e)
        {
            SetCurrentPage((ConfigFormPage)((ToolStripButton)sender).Tag);
        }

        private void SetCurrentPage(ConfigFormPage page)
        {
            _configFormViewModel.SetPage(page);
        }

        private void ConfigureForm_Load(object sender, EventArgs e)
        {
            _configFormViewModel.PropertyChanged += (s, a) =>
            {
                if (a.PropertyName != nameof(_configFormViewModel.CurrentPage))
                {
                    return;
                }

                ShowPage(_configFormViewModel.CurrentPage);
            };
        }

        private void ShowPage(ConfigFormPage page)
        {
            SuspendLayout();

            SetPageEnabled(optionsPage, optionsButton, page == ConfigFormPage.Options);
            SetPageEnabled(fileStructurePage, filesButton, page == ConfigFormPage.Files);
            SetPageEnabled(folderStructurePage, foldersButton, page == ConfigFormPage.Folders);
            SetPageEnabled(rulesPage, rulesButton, page == ConfigFormPage.Rules);

            ResumeLayout();
        }

        private static void SetPageEnabled(
            System.Windows.Forms.Control page,
            ToolStripButton button,
            bool enabled
        )
        {
            if (enabled)
            {
                page.Show();
            }
            else
            {
                page.Hide();
            }

            button.Checked = enabled;
        }

        private void ConfigureForm_ResizeBegin(object sender, EventArgs e)
        {
            SuspendLayout();
        }

        private void ConfigureForm_ResizeEnd(object sender, EventArgs e)
        {
            ResumeLayout();
        }
    }
}
