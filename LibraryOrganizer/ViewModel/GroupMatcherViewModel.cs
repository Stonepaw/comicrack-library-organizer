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
            Matchers.Add(
                new BookFieldStringMatcherViewModel(this, ComicBookField.ComicBookField.Series)
            );
        }

        public void AddBookFieldMatcher(IBookFieldMatcher matcher)
        {
            switch (matcher)
            {
                case IBookFieldIntMatcher bookFieldIntMatcher:
                    Matchers.Add(new BookFieldIntMatcherViewModel(this, bookFieldIntMatcher));
                    break;
                case IBookFieldStringMatcher bookFieldMatcher:
                    Matchers.Add(new BookFieldStringMatcherViewModel(this, bookFieldMatcher));
                    break;
                case IBookFieldYesNoMatcher bookFieldYesNoMatcher:
                    Matchers.Add(new BookFieldYesNoMatcherViewModel(this, bookFieldYesNoMatcher));
                    break;
            }
        }

        public void AddGroup()
        {
            Matchers.Add(new GroupMatcherViewModel(this));
        }

        public void AddGroup(IGroupMatcher groupMatcher)
        {
            Matchers.Add(new GroupMatcherViewModel(this, groupMatcher));
        }

        public void ChangeFieldType(IBookFieldMatcherViewModel matcher, IComicBookField field)
        {
            IBookFieldMatcherViewModel replacement;

            switch (field)
            {
                case IComicBookIntField intField:
                    replacement = new BookFieldIntMatcherViewModel(this, intField);
                    break;
                case IComicBookStringField stringField:
                    replacement = new BookFieldStringMatcherViewModel(this, stringField);
                    break;
                case IComicBookYesNoField yesNoField:
                    replacement = new BookFieldYesNoMatcherViewModel(this, yesNoField);
                    break;
                default:
                    throw new NotImplementedException();
            }

            Matchers[Matchers.IndexOf(matcher)] = replacement;
        }

        public void Delete()
        {
            _parent?.Remove(this);
        }

        public void Remove(IMatcherViewModel matcher)
        {
            Matchers.Remove(matcher);
        }
    }
}
