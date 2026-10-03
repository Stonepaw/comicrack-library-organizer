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
        private readonly BindingSource _bindingSource;

        private readonly IGroupMatcherViewModel _matcher;

        public new int Width
        {
            get => base.Width;
            set
            {
                SuspendLayout();
                base.Width = value;
                base.MaximumSize = new Size(value, int.MaxValue);
                base.MinimumSize = new Size(value, 0);

                matchersPanel.SuspendLayout();

                int controlWidth =
                    value
                    - Margin.Horizontal
                    - matchersPanel.Margin.Horizontal
                    - matchersPanel.Padding.Horizontal;

                foreach (System.Windows.Forms.Control control in matchersPanel.Controls)
                {
                    int newWidth = controlWidth - control.Margin.Horizontal;

                    if (control.Width == newWidth)
                    {
                        continue;
                    }

                    switch (control)
                    {
                        case MatcherGroupControl groupControl:
                            groupControl.Width = newWidth;
                            break;
                        case MatcherRuleControl ruleControl:
                            ruleControl.Width = newWidth;
                            break;
                        default:
                            control.Width = newWidth;
                            break;
                    }
                }

                matchersPanel.ResumeLayout();
                ResumeLayout();
            }
        }

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
            matchersPanel.RebuildMatcherControls(
                _matcher.Matchers,
                matchersPanel.ClientSize.Width - matchersPanel.Padding.Horizontal
            );
        }

        private void Matchers_ListChanged(object sender, ListChangedEventArgs e)
        {
            MatcherGroupControlUtils.MatcherControlsListChanged(
                e,
                matchersPanel,
                _matcher.Matchers,
                matchersPanel.ClientSize.Width - matchersPanel.Padding.Horizontal
            );
        }

        private void matchActionsButton_Click(object sender, EventArgs e)
        {
            matcherGroupActions.Show(actionsButton, new Point(0, actionsButton.Height));
        }

        private void addMatcherAction_Click(object sender, EventArgs e)
        {
            _matcher.AddBookFieldMatcher();
        }

        private void addGroupAction_Click(object sender, EventArgs e)
        {
            _matcher.AddGroup();
        }
    }
}
