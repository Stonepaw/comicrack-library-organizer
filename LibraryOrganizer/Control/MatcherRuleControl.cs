using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace LibraryOrganizer.Control
{
    public partial class MatcherRuleControl : UserControl
    {
        public new int Width
        {
            get => base.Width;
            set
            {
                Debug.WriteLine($"Setting matcher width to {value}");

                base.Width = value;
                MaximumSize = new Size(value, int.MaxValue);
                MinimumSize = new Size(value, 0);
            }
        }

        public MatcherRuleControl()
        {
            InitializeComponent();

            comboBox1.DataSource = ComicBookField.ComicBookField.Fields;
            comboBox1.DisplayMember = "Label";
        }
    }
}
