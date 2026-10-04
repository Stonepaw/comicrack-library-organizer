using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
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

            Control control = panel.Controls[fromIndex];

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
            IMatcherViewModel iMatcherViewModel,
            int width,
            int index = -1
        )
        {
            switch (iMatcherViewModel)
            {
                case IGroupMatcherViewModel groupMatcherViewModel:
                    MatcherGroupControl groupControl = new MatcherGroupControl(
                        groupMatcherViewModel
                    );
                    groupControl.AutoSize = true;
                    groupControl.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                    groupControl.Width = width;

                    AddMatcherControlToPanel(panel, groupControl, width, index);
                    break;
                case BookFieldMatcherViewModel matcherViewModel:
                    MatcherControl stringMatcherControl = new MatcherControl(matcherViewModel);
                    stringMatcherControl.Width = width;
                    AddMatcherControlToPanel(panel, stringMatcherControl, width, index);
                    break;
            }
        }

        private static void AddMatcherControlToPanel(
            Panel panel,
            Control control,
            int width,
            int index
        )
        {
            control.SuspendLayout();
            control.Width = width;
            control.MinimumSize = new Size(width, 0);
            control.MaximumSize = new Size(width, int.MaxValue);
            panel.Controls.Add(control);

            if (index >= 0)
            {
                panel.Controls.SetChildIndex(control, index);
            }

            control.ResumeLayout();
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
