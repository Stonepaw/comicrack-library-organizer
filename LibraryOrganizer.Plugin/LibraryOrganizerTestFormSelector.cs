using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
            new ConfigureForm(new Data.Profile()).Show(this);
        }
    }
}
