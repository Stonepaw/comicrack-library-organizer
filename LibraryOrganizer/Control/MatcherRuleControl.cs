using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LibraryOrganizer.Control
{
    public partial class MatcherRuleControl : UserControl
    {
        public MatcherRuleControl()
        {
            InitializeComponent();

            comboBox1.DataSource = ComicBookField.ComicBookField.Fields;
            comboBox1.DisplayMember = "Label";
        }
    }
}
