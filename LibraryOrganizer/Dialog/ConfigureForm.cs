using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using LibraryOrganizer.Data;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Dialog
{
    public partial class ConfigureForm : Form
    {
        private readonly List<ProfileViewModel> _profiles;

        private readonly ConfigFormViewModel _configFormViewModel = new ConfigFormViewModel();

        private bool _fileButtonEnabled = true;
        private bool _folderButtonEnabled = true;

        /// <summary>
        /// This is used to databind the file toolstrip button since toolstrip buttons don't have data binding.
        /// </summary>
        public bool FileButtonEnabled
        {
            get => _fileButtonEnabled;
            set
            {
                _fileButtonEnabled = value;
                filesButton.Enabled = value;
            }
        }

        /// <summary>
        /// This is used to databind the folder toolstrip button since toolstrip buttons don't have data binding.
        /// </summary>
        public bool FolderButtonEnabled
        {
            get => _folderButtonEnabled;
            set
            {
                _folderButtonEnabled = value;
                foldersButton.Enabled = value;
            }
        }

        /// <summary>
        /// This is used to simplify data binding which page is visible.
        /// </summary>
        public ConfigFormPage CurrentPage
        {
            get => _configFormViewModel.CurrentPage;
            set => ShowPage(value);
        }

        private ProfileViewModel CurrentProfileViewModel =>
            (ProfileViewModel)profileBindingSource.Current;

        public ConfigureForm()
            : this(new[] { new Profile() }) { }

        public ConfigureForm(IEnumerable<Profile> profiles)
        {
            InitializeComponent();

            SuspendLayout();

            _profiles = profiles.Select(profile => new ProfileViewModel(profile)).ToList();

            profileBindingSource.DataSource = _profiles;
            configFormViewModelBindingSource.DataSource = _configFormViewModel;

            removeEmptyFolderExclusions.SelectedIndex = -1;
            overviewButton.Tag = ConfigFormPage.Overview;
            filesButton.Tag = ConfigFormPage.Files;
            foldersButton.Tag = ConfigFormPage.Folders;
            rulesButton.Tag = ConfigFormPage.Rules;
            optionsButton.Tag = ConfigFormPage.Options;

            if (profileSelector.ComboBox != null)
            {
                profileSelector.ComboBox.DataSource = profileBindingSource;
                profileSelector.ComboBox.DisplayMember = "Name";
            }

            // Some things can't be data bound directly. We can instead data-bind our form properties and use them as
            // updater functions to update form properties.
            DataBindings.Add(
                new Binding("FileButtonEnabled", profileBindingSource, "UseFileNaming")
            );
            DataBindings.Add(
                new Binding("FolderButtonEnabled", profileBindingSource, "UseFolderOrganization")
            );
            DataBindings.Add("CurrentPage", _configFormViewModel, "CurrentPage");
            FileButtonEnabled = CurrentProfileViewModel.UseFileNaming;
            FolderButtonEnabled = CurrentProfileViewModel.UseFolderOrganization;
            SetCurrentPage(ConfigFormPage.Overview);
            ShowPage(ConfigFormPage.Overview);

            profileBindingSource.CurrentChanged += (s, e) =>
            {
                if (CurrentProfileViewModel != null)
                {
                    overviewPage.ProfileViewModel = CurrentProfileViewModel;
                }
            };
            overviewPage.ProfileViewModel = CurrentProfileViewModel;

            ResumeLayout();
        }

        private void failOperationOnEmptyValueDestinationFolderBrowse_Click(
            object sender,
            EventArgs e
        )
        {
            string folder = SelectFolder();

            if (folder != null)
            {
                CurrentProfileViewModel.FailOperationOnEmptyValueDestinationFolder = folder;
            }
        }

        private void addEmptyFolderExclusion_Click(object sender, EventArgs e)
        {
            string folder = SelectFolder();

            if (folder != null)
            {
                removeEmptyFoldersExclusionsBindingSource.Add(folder);
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
                CurrentProfileViewModel.IllegalCharacterReplacements
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

        private void ConfigureForm_Load(object sender, EventArgs e) { }

        private void ShowPage(ConfigFormPage page)
        {
            SuspendLayout();

            SetPageEnabled(overviewPage, overviewButton, page == ConfigFormPage.Overview);
            SetPageEnabled(optionsPage, optionsButton, page == ConfigFormPage.Options);
            SetPageEnabled(fileStructurePage, filesButton, page == ConfigFormPage.Files);
            SetPageEnabled(folderStructurePage, foldersButton, page == ConfigFormPage.Folders);
            SetPageEnabled(rulesPage, rulesButton, page == ConfigFormPage.Rules);

            ResumeLayout();
        }

        private static void SetPageEnabled(Control page, ToolStripButton button, bool enabled)
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

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddProfileDialog dialog = new AddProfileDialog();

            DialogResult result = dialog.ShowDialog(this);

            if (result != DialogResult.OK)
            {
                return;
            }

            ProfileViewModel profile = new ProfileViewModel(new Profile())
            {
                Name = dialog.ProfileName,
            };

            profileBindingSource.Add(profile);
            profileBindingSource.Position = profileBindingSource.Count;
        }

        private string SelectFolder()
        {
            var openFolderDialog = new FolderBrowserDialog();

            if (
                openFolderDialog.ShowDialog(this) == DialogResult.OK
                && openFolderDialog.SelectedPath != null
            )
            {
                return openFolderDialog.SelectedPath;
            }

            return null;
        }
    }
}
