using System;
using System.ComponentModel;
using LibraryOrganizer.ComicBookField;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class GroupMatcherViewModel : IGroupMatcherViewModel
    {
        private readonly IGroupMatcherViewModel _parent;

        public BindingList<IMatcherViewModel> Matchers { get; } =
            new BindingList<IMatcherViewModel>();

        public GroupMatcherMode Mode { get; set; }

        public bool Negated { get; set; }

        public GroupMatcherViewModel(IGroupMatcherViewModel parent)
        {
            _parent = parent;
        }

        public GroupMatcherViewModel(IGroupMatcherViewModel parent, IGroupMatcher groupMatcher)
        {
            _parent = parent;

            foreach (IMatcher matcher in groupMatcher.Matchers)
            {
                switch (matcher)
                {
                    case IBookFieldMatcher bookFieldMatcher:
                        AddBookFieldMatcher(bookFieldMatcher);
                        break;
                    case IGroupMatcher subGroupMatcher:
                        AddGroup(subGroupMatcher);
                        break;
                }
            }
        }

        public void AddBookFieldMatcher()
        {
            AddBookFieldMatcherAt(Matchers.Count, null);
        }

        public void AddBookFieldMatcher(IBookFieldMatcher matcher)
        {
            AddBookFieldMatcherAt(Matchers.Count, matcher);
        }

        public void AddBookFieldMatcherAfter(IMatcherViewModel after)
        {
            AddBookFieldMatcherAt(Matchers.IndexOf(after) + 1, null);
        }

        public void AddGroup()
        {
            AddGroupAt(Matchers.Count, null);
        }

        public void AddGroup(IGroupMatcher groupMatcher)
        {
            AddGroupAt(Matchers.Count, groupMatcher);
        }

        public void AddGroupAfter(IMatcherViewModel after)
        {
            AddGroupAt(Matchers.IndexOf(after) + 1, null);
        }

        public void Delete()
        {
            _parent?.Remove(this);
        }

        public void Remove(IMatcherViewModel matcher)
        {
            Matchers.Remove(matcher);
        }

        private void AddGroupAt(int index, IGroupMatcher matcher)
        {
            AddMatcherAt(
                index,
                matcher != null
                    ? new GroupMatcherViewModel(this, matcher)
                    : new GroupMatcherViewModel(this)
            );
        }

        private void AddBookFieldMatcherAt(int index, IBookFieldMatcher matcher)
        {
            AddMatcherAt(
                index,
                matcher != null
                    ? new BookFieldMatcherViewModel(this, matcher)
                    : new BookFieldMatcherViewModel(this)
            );
        }

        private void AddMatcherAt(int index, IMatcherViewModel viewModel)
        {
            if (index >= 0 && index < Matchers.Count)
            {
                Matchers.Insert(index, viewModel);
            }
            else
            {
                Matchers.Add(viewModel);
            }
        }
    }
}
