using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class EmptyValuesConfigControl : UserControl
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

        public EmptyValuesConfigControl()
        {
            InitializeComponent();
        }

        private void failOperationOnEmptyValueDestinationFolderBrowse_Click(
            object sender,
            EventArgs e
        )
        {
            string folder = FolderBrowserUtil.PromptForFolderPath(this);

            if (folder != null && _profileViewModel != null)
            {
                _profileViewModel.FailOperationOnEmptyValueDestinationFolder = folder;
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
