using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Control
{
    internal partial class ProfileMatcherGroupControl : UserControl
    {
        private IGroupMatcherViewModel _matcher;

        public IGroupMatcherViewModel Matcher
        {
            get => _matcher;
            set
            {
                if (_matcher != null)
                {
                    _matcher.Matchers.ListChanged -= MatchersOnListChanged;
                }

                _matcher = value;

                if (value == null)
                {
                    return;
                }

                value.Matchers.ListChanged += MatchersOnListChanged;
                matchersPanel.RebuildMatcherControls(
                    value.Matchers,
                    matchersPanel.Size.Width
                        - matchersPanel.Margin.Horizontal
                        - SystemInformation.VerticalScrollBarWidth
                        - 10
                );
            }
        }

        public ProfileMatcherGroupControl()
        {
            InitializeComponent();

            matcherGroupActionMenuButton.Left =
                configPanel.ClientSize.Width
                - matcherGroupActionMenuButton.Width
                - SystemInformation.VerticalScrollBarWidth
                - 9;
        }

        private void MatchersOnListChanged(object sender, ListChangedEventArgs e)
        {
            MatcherGroupControlUtils.MatcherControlsListChanged(
                e,
                matchersPanel,
                Matcher.Matchers,
                matchersPanel.Size.Width
                    - matchersPanel.Margin.Horizontal
                    - SystemInformation.VerticalScrollBarWidth
                    - 10
            );
        }

        private void addGroup_Click(object sender, System.EventArgs e)
        {
            Matcher?.AddGroup();
        }

        private void addRule_Click(object sender, System.EventArgs e)
        {
            Matcher?.AddBookFieldMatcher();
        }

        private void matcherGroupActionMenuButton_Click(object sender, System.EventArgs e)
        {
            matcherGroupActionMenu.Show(
                matcherGroupActionMenuButton,
                new Point(0, matcherGroupActionMenuButton.Height)
            );
        }

        private void ProfileMatcherGroupControl_Resize(object sender, System.EventArgs e)
        {
            SuspendLayout();

            int panelWidth = Width - Padding.Horizontal - matchersPanel.Margin.Horizontal;

            int matcherControlWidth = panelWidth - SystemInformation.VerticalScrollBarWidth - 10;

            flowLayoutPanel1.SuspendLayout();

            foreach (System.Windows.Forms.Control control in matchersPanel.Controls)
            {
                int controlWidth = control.Width;

                if (controlWidth == matcherControlWidth)
                {
                    continue;
                }

                switch (control)
                {
                    case MatcherGroupControl groupControl:
                        groupControl.Width = matcherControlWidth;
                        break;
                    case StringMatcherControl matcherRuleControl:
                        matcherRuleControl.Width = matcherControlWidth;
                        break;
                    default:
                        control.Width = matcherControlWidth;
                        control.MinimumSize = new Size(matcherControlWidth, 0);
                        control.MaximumSize = new Size(matcherControlWidth, int.MaxValue);
                        break;
                }
            }

            flowLayoutPanel1.ResumeLayout();
            ResumeLayout();
        }
    }
}
