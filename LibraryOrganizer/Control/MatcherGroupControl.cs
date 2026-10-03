using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Control
{
    internal partial class MatcherGroupControl : UserControl
    {
        private BindingSource _bindingSource;

        private IGroupMatcherViewModel _matcher;

        public MatcherGroupControl()
        {
            _bindingSource = new BindingSource();

            InitializeComponent();
        }

        public MatcherGroupControl(IGroupMatcherViewModel matcher)
        {
            _bindingSource = new BindingSource();
            _bindingSource.DataSource = matcher;

            InitializeComponent();

            _matcher = matcher;
            matcher.Matchers.ListChanged += Matchers_ListChanged;
            flowLayoutPanel1.RebuildMatcherControls(
                _matcher.Matchers,
                flowLayoutPanel1.ClientSize.Width - flowLayoutPanel1.Padding.Horizontal
            );
        }

        private void Matchers_ListChanged(object sender, ListChangedEventArgs e)
        {
            MatcherGroupControlUtils.MatcherControlsListChanged(
                e,
                flowLayoutPanel1,
                _matcher.Matchers,
                flowLayoutPanel1.ClientSize.Width - flowLayoutPanel1.Padding.Horizontal
            );
        }

        private void button1_Click(object sender, EventArgs e)
        {
            contextMenuStrip1.Show(button1, new Point(0, button1.Height));
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            _matcher.AddBookFieldMatcher();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            _matcher.AddGroup();
        }

        private void flowLayoutPanel1_ClientSizeChanged(object sender, EventArgs e)
        {
            Debug.WriteLine($"Group client size changed {flowLayoutPanel1.ClientSize.Width}");

            //flowLayoutPanel1.MakeAllControlsFullWidth();
        }

        private void flowLayoutPanel1_SizeChanged(object sender, EventArgs e)
        {
            Debug.WriteLine($"Group size changed {flowLayoutPanel1.ClientSize.Width}");
        }

        private void flowLayoutPanel1_Layout(object sender, LayoutEventArgs e)
        {
            Debug.WriteLine($"Group layout changed {flowLayoutPanel1.ClientSize.Width}");
        }
    }
}
