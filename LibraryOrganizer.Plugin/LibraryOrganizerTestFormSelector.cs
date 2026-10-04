using System;
using System.Windows.Forms;
using LibraryOrganizer.Data;
using LibraryOrganizer.Dialog;

namespace LibraryOrganizer.Plugin
{
    public partial class LibraryOrganizerTestFormSelector : Form
    {
        public LibraryOrganizerTestFormSelector()
        {
            InitializeComponent();
        }

        private void openConfigureForm_Click(object sender, EventArgs e)
        {
            Profile profile = new Profile();
            profile.Name = "Test";

            new ConfigureForm(new[] { profile }).Show(this);
        }
    }
}
