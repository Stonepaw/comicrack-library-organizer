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
            // For whatever reason the bindings don't work when created in the designer so we have to add them here.
            optionsPage.DataBindings.Add(
                "Visible",
                configFormViewModelBindingSource,
                "OptionsPageEnabled"
            );
            optionsPage.DataBindings.Add(
                "Enabled",
                configFormViewModelBindingSource,
                "OptionsPageEnabled"
            );
            rulesPage.DataBindings.Add(
                "Visible",
                configFormViewModelBindingSource,
                "RulesPageEnabled"
            );
            rulesPage.DataBindings.Add(
                "Enabled",
                configFormViewModelBindingSource,
                "RulesPageEnabled"
            );
            folderStructurePage.DataBindings.Add(
                "Visible",
                configFormViewModelBindingSource,
                "FoldersPageEnabled"
            );
            folderStructurePage.DataBindings.Add(
                "Enabled",
                configFormViewModelBindingSource,
                "FoldersPageEnabled"
            );
            fileStructurePage.DataBindings.Add(
                "Visible",
                configFormViewModelBindingSource,
                "FilesPageEnabled"
            );
            fileStructurePage.DataBindings.Add(
                "Enabled",
                configFormViewModelBindingSource,
                "FilesPageEnabled"
            );

            // ToolStripButton doesn't have bindings, so we just hook them up here.
            _configFormViewModel.PropertyChanged += (o, args) =>
            {
                switch (args.PropertyName)
                {
                    case nameof(ConfigFormViewModel.FilesPageEnabled):
                        filesButton.Checked = _configFormViewModel.FilesPageEnabled;
                        break;
                    case nameof(ConfigFormViewModel.FoldersPageEnabled):
                        foldersButton.Checked = _configFormViewModel.FoldersPageEnabled;
                        break;
                    case nameof(ConfigFormViewModel.OptionsPageEnabled):
                        optionsButton.Checked = _configFormViewModel.OptionsPageEnabled;
                        break;
                    case nameof(ConfigFormViewModel.RulesPageEnabled):
                        rulesButton.Checked = _configFormViewModel.RulesPageEnabled;
                        break;
                    case nameof(ConfigFormViewModel.OverviewPageEnabled):
                        overviewButton.Checked = _configFormViewModel.OverviewPageEnabled;
                        break;
                }
            };
        }
    }
}
