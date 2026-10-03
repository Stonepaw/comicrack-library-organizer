using System.ComponentModel;
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
                    matchersPanel.Size.Width - SystemInformation.VerticalScrollBarWidth - 10
                );
            }
        }

        public ProfileMatcherGroupControl()
        {
            InitializeComponent();

            // TODO: Figure out a nicer solution for this
            matcherGroupActionMenuButton.Location = new Point(
                configPanel.ClientSize.Width
                    - configPanel.Padding.Left
                    - matcherGroupActionMenuButton.Width
                    - SystemInformation.VerticalScrollBarWidth
                    - 3,
                3
            );
        }

        private void MatchersOnListChanged(object sender, ListChangedEventArgs e)
        {
            MatcherGroupControlUtils.MatcherControlsListChanged(
                e,
                matchersPanel,
                Matcher.Matchers,
                matchersPanel.Size.Width - SystemInformation.VerticalScrollBarWidth - 10
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
    }
}
