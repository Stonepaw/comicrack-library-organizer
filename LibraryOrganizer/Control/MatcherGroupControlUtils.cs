using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Control
{
    /// <summary>
    /// This handles shared actions for matcher groups to handle changes to the list of matchers.
    ///
    /// This is shared since we have two types of matcher groups: one for the top level profile
    /// group and other for nested groups.
    /// </summary>
    internal static class MatcherGroupControlUtils
    {
        public static void MatcherControlsListChanged(
            ListChangedEventArgs e,
            Panel panel,
            IReadOnlyList<IMatcherViewModel> matchers,
            int width
        )
        {
            switch (e.ListChangedType)
            {
                case ListChangedType.ItemAdded:
                    panel.AddMatcher(matchers[e.NewIndex], width, e.NewIndex);
                    break;
                case ListChangedType.ItemChanged:
                    panel.ChangeMatcher(matchers[e.NewIndex], width, e.NewIndex);
                    break;
                case ListChangedType.ItemDeleted:
                    panel.RemoveMatcher(e.NewIndex);
                    break;
                case ListChangedType.ItemMoved:
                    panel.MoveMatcher(e.OldIndex, e.NewIndex);
                    break;
                case ListChangedType.Reset:
                    panel.RebuildMatcherControls(matchers, width);
                    break;
            }
        }

        public static void RebuildMatcherControls(
            this Panel panel,
            IEnumerable<IMatcherViewModel> matchers,
            int width
        )
        {
            panel.SuspendLayout();
            panel.Controls.Clear();

            foreach (IMatcherViewModel matcherViewModel in matchers)
            {
                panel.AddMatcherControl(matcherViewModel, width);
            }

            panel.AutoTabIndex();
            panel.ResumeLayout();
        }

        public static void MakeAllControlsFullWidth(this Panel panel)
        {
            panel.SuspendLayout();

            var width = panel.ClientSize.Width - panel.Padding.Horizontal;

            foreach (System.Windows.Forms.Control control in panel.Controls)
            {
                control.Width = width - control.Margin.Horizontal;

                if (control is MatcherGroupControl)
                {
                    control.PerformLayout();
                }
            }

            panel.ResumeLayout(true);
        }

        private static void AddMatcher(
            this Panel panel,
            IMatcherViewModel matcher,
            int width,
            int index
        )
        {
            panel.SuspendLayout();
            panel.AddMatcherControl(matcher, width, index);
            panel.AutoTabIndex();
            panel.ResumeLayout();
        }

        private static void ChangeMatcher(
            this Panel panel,
            IMatcherViewModel matcherViewModel,
            int width,
            int index
        )
        {
            if (index >= panel.Controls.Count)
            {
                return;
            }

            panel.SuspendLayout();
            panel.Controls.RemoveAt(index);
            panel.AddMatcherControl(matcherViewModel, width, index);
            panel.AutoTabIndex();
            panel.ResumeLayout();
        }

        private static void MoveMatcher(this Panel panel, int fromIndex, int toIndex)
        {
            if (fromIndex >= panel.Controls.Count)
            {
                return;
            }

            System.Windows.Forms.Control control = panel.Controls[fromIndex];

            if (control == null)
            {
                return;
            }

            panel.SuspendLayout();
            panel.Controls.SetChildIndex(control, toIndex);
            panel.AutoTabIndex();
            panel.ResumeLayout();
        }

        private static void RemoveMatcher(this Panel panel, int index)
        {
            panel.SuspendLayout();
            panel.Controls.RemoveAt(index);
            panel.AutoTabIndex();
            panel.ResumeLayout();
        }

        private static void AddMatcherControl(
            this Panel panel,
            IMatcherViewModel matcherViewModel,
            int width,
            int index = -1
        )
        {
            switch (matcherViewModel)
            {
                case IGroupMatcherViewModel groupMatcherViewModel:
                    MatcherGroupControl groupControl = new MatcherGroupControl(
                        groupMatcherViewModel
                    );
                    groupControl.AutoSize = true;
                    groupControl.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                    AddMatcherControlToPanel(panel, groupControl, width, index);
                    break;
                case IBookFieldMatcherViewModel bookFieldMatcherViewModel:
                    AddMatcherControlToPanel(panel, new MatcherRuleControl(), width, index);
                    break;
            }
        }

        private static void AddMatcherControlToPanel(
            Panel panel,
            System.Windows.Forms.Control control,
            int width,
            int index
        )
        {
            control.Width = width;
            panel.Controls.Add(control);

            if (index >= 0)
            {
                panel.Controls.SetChildIndex(control, index);
            }
        }

        private static void AutoTabIndex(this Panel panel)
        {
            for (var i = 0; i < panel.Controls.Count; i++)
            {
                panel.Controls[i].TabIndex = i;
            }
        }
    }
}
