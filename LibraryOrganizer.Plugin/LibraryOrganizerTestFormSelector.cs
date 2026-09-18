using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            var configureForm = new ConfigureForm();
            configureForm.Show(this);
        }
    }
}
