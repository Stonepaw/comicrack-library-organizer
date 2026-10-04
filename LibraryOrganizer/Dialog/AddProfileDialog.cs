using System;
using System.Windows.Forms;

namespace LibraryOrganizer.Dialog
{
    public partial class AddProfileDialog : Form
    {
        public string ProfileName => profileName.Text;

        public AddProfileDialog()
        {
            InitializeComponent();
        }

        private void okay_Click(object sender, EventArgs e)
        {
            if (profileName.Text.Length <= 0)
            {
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
